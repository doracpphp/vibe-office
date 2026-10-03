"""エージェントループとツール選択のテスト（API は呼ばずにフェイクのクライアントを使う）"""
import json
from types import SimpleNamespace

import pytest
from openai.types.chat import ChatCompletionMessage

import agent


@pytest.fixture(autouse=True)
def workdir(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    return tmp_path


# ── ツール選択 ────────────────────────────────────────────────────────────────

def _names(tools):
    return {t["name"] for t in tools}


def test_select_tools_ignores_system_prompt():
    history = [{"role": "system", "content": agent.SYSTEM_PROMPT},
               {"role": "user", "content": "sales.xlsxを開いてA1を読んで"}]
    tools, openai_tools = agent._select_tools(history)
    assert "read_cell" in _names(tools)
    assert "read_document" not in _names(tools)
    assert len(openai_tools) == len(tools)


def test_select_tools_word_boundaries():
    def selected(text):
        return _names(agent._select_tools([{"role": "user", "content": text}])[0])

    # "password" は Word と判定しない（判定不能 → 全ツール）
    assert "read_cell" in selected("password を変更して")
    assert "read_cell" not in selected("report.docxを開いて")
    assert "parse_markdown" in selected("README.md を Word に変換して")


# ── フェイククライアント ──────────────────────────────────────────────────────

class FakeAnthropic:
    def __init__(self, responses):
        self.responses = list(responses)
        self.requests = []
        self.messages = self

    def create(self, **kwargs):
        self.requests.append(kwargs)
        item = self.responses.pop(0)
        if isinstance(item, Exception):
            raise item
        return item


def _text(text):
    return SimpleNamespace(type="text", text=text)


def _tool_use(id_, name, input_):
    return SimpleNamespace(type="tool_use", id=id_, name=name, input=input_)


def _anthropic_agent(responses):
    a = agent._AnthropicAgent.__new__(agent._AnthropicAgent)
    a._client = FakeAnthropic(responses)
    a._model = "test"
    agent._BaseAgent.__init__(a)
    return a


def test_anthropic_tool_loop(workdir):
    a = _anthropic_agent([
        SimpleNamespace(stop_reason="tool_use", content=[
            _tool_use("t1", "write_cell", {"file_path": "a.xlsx", "cell_address": "A1", "value": 5})]),
        SimpleNamespace(stop_reason="end_turn", content=[_text("書き込みました")]),
    ])
    assert a.chat("a.xlsx の A1 に 5") == "書き込みました"
    assert (workdir / "a.xlsx").exists()

    req = a._client.requests[1]
    # システムプロンプトと最後のメッセージにキャッシュ指定が付く
    assert req["system"][0]["cache_control"] == {"type": "ephemeral"}
    assert req["messages"][-1]["content"][-1]["cache_control"] == {"type": "ephemeral"}
    # 履歴そのものにはキャッシュ指定を残さない
    assert "cache_control" not in a._history[2]["content"][-1]


def test_anthropic_history_rolls_back_on_error():
    a = _anthropic_agent([
        SimpleNamespace(stop_reason="tool_use", content=[
            _tool_use("t1", "list_sheets", {"file_path": "missing.xlsx"})]),
        RuntimeError("API 障害"),
    ])
    with pytest.raises(RuntimeError):
        a.chat("missing.xlsx のシート一覧")
    # tool_use だけが残った壊れた履歴にならない
    assert a._history == []


def test_anthropic_max_tokens_with_partial_tool_use():
    a = _anthropic_agent([
        SimpleNamespace(stop_reason="max_tokens", content=[
            _text("途中まで"), _tool_use("t1", "write_range", {"file_path": "a.xlsx"})]),
    ])
    reply = a.chat("大量データを書いて")
    assert reply.startswith("途中まで")
    assert "上限" in reply
    # tool_result の無い tool_use を履歴に残さない
    assert a._history[-1] == {"role": "assistant", "content": "途中まで"}


def test_anthropic_tool_round_limit(monkeypatch):
    monkeypatch.setattr(agent, "_MAX_TOOL_ROUNDS", 2)
    loop = SimpleNamespace(stop_reason="tool_use", content=[
        _tool_use("t", "list_sheets", {"file_path": "x.xlsx"})])
    a = _anthropic_agent([loop, loop, loop])
    assert "中断" in a.chat("x.xlsx")
    assert len(a._client.requests) == 2


class FakeOpenAI:
    def __init__(self, messages):
        self.messages = list(messages)
        self.requests = []
        self.chat = SimpleNamespace(completions=self)

    def create(self, **kwargs):
        self.requests.append(json.loads(json.dumps(kwargs["messages"])))
        finish, message = self.messages.pop(0)
        return SimpleNamespace(choices=[SimpleNamespace(
            finish_reason=finish, message=ChatCompletionMessage.model_validate(message))])


def _openai_agent(messages):
    a = agent._OpenAICompatAgent.__new__(agent._OpenAICompatAgent)
    a._client = FakeOpenAI(messages)
    a._model = "test"
    agent._BaseAgent.__init__(a)
    return a


def _tool_call(id_, name, arguments):
    return {"id": id_, "type": "function", "function": {"name": name, "arguments": arguments}}


def test_openai_tool_calls_with_stop_finish_reason(workdir):
    # Gemini / Ollama は tool_calls があっても finish_reason="stop" を返すことがある
    a = _openai_agent([
        ("stop", {"role": "assistant", "content": None, "tool_calls": [
            _tool_call("c1", "write_cell", '{"file_path": "a.xlsx", "cell_address": "A1", "value": 1}')]}),
        ("stop", {"role": "assistant", "content": "完了"}),
    ])
    assert a.chat("a.xlsx の A1 に 1") == "完了"
    assert (workdir / "a.xlsx").exists()
    assert a._history[-2]["role"] == "tool"


def test_openai_invalid_arguments_are_reported_to_model():
    a = _openai_agent([
        ("tool_calls", {"role": "assistant", "content": None, "tool_calls": [
            _tool_call("c1", "read_cell", '{"file_path": ')]}),
        ("stop", {"role": "assistant", "content": "やり直します"}),
    ])
    assert a.chat("a.xlsx") == "やり直します"
    tool_msg = a._client.requests[1][-1]
    assert tool_msg["role"] == "tool"
    assert json.loads(tool_msg["content"])["success"] is False


def test_reset_keeps_system_prompt():
    a = _openai_agent([("stop", {"role": "assistant", "content": "hi"})])
    a.chat("hello")
    a.reset()
    assert a._history == [{"role": "system", "content": agent.SYSTEM_PROMPT}]
