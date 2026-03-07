# vibe-office

> チャットベースで Excel・Word ファイルを操作する AI エージェント — Claude / OpenRouter / Ollama 対応

[English README](README.md)

## 機能

- **Excel** — セル・範囲・数式の読み書き、シート管理、書式設定
- **Word** — ドキュメントの読み取り、段落の挿入・削除・置換、見出し・テーブル・画像の追加
- **マルチプロバイダー** — Anthropic Claude・Google Gemini・OpenRouter（クラウド）・Ollama（ローカル・オフライン）に対応
- **スマートツール選択** — 会話の文脈から Excel/Word を自動判定し、必要なツールだけをモデルに渡してトークン消費を削減
- **スピナー表示** — エージェント処理中は実行中のツール名をアニメーション表示

## 必要環境

- Python 3.11 以上
- [uv](https://docs.astral.sh/uv/)
- 使用するプロバイダーの API キー

## インストール

```bash
git clone <repo-url>
cd vibe-office
uv sync
```

## 設定

### .env ファイルを使う方法（推奨）

`.env.example` をコピーして `.env` を作成し、必要なAPIキーを設定してください。`.env` は `.gitignore` に含まれており、**Git にはコミットされません**。

```bash
cp .env.example .env
```

`.env` の内容例:

```
ANTHROPIC_API_KEY=sk-ant-...
GEMINI_API_KEY=AIza...
OPENROUTER_API_KEY=sk-or-...
```

起動時に読み込むには:

```bash
# bash / zsh
source .env && uv run python main.py

# または direnv を使えば自動で読み込まれます
```

### 環境変数を直接設定する方法

```bash
# Anthropic（デフォルト）
export ANTHROPIC_API_KEY="sk-ant-..."

# Google Gemini
# https://aistudio.google.com/app/apikey で取得
export GEMINI_API_KEY="AIza..."

# OpenRouter
export OPENROUTER_API_KEY="sk-or-..."

# Ollama — キー不要（ローカルで動作）
```

> **注意**: APIキーはコードに直接書かないでください。環境変数または `.env` ファイルで管理してください。

## 起動方法

```bash
# Anthropic Claude（デフォルト）
uv run python main.py

# OpenRouter
uv run python main.py --provider openrouter

# OpenRouter でモデルを指定
uv run python main.py --provider openrouter --model google/gemini-flash-1.5

# Google Gemini
uv run python main.py --provider gemini

# Google Gemini でモデルを指定
uv run python main.py --provider gemini --model gemini-2.5-pro

# Ollama（ローカル）
uv run python main.py --provider ollama

# Ollama でモデルやエンドポイントを指定
uv run python main.py --provider ollama --model qwen3.5
uv run python main.py --provider ollama --base-url http://192.168.1.10:11434/v1
```

フラグの代わりに環境変数でも指定できます:

```bash
export EXCEL_AGENT_PROVIDER=ollama
export EXCEL_AGENT_MODEL=qwen2.5:7b
uv run python main.py
```

## 使用例

```
You> sales.xlsx を開いて Sheet1 の内容を見せて
You> A1 に「売上」、B1 に 250000 を書いて
You> C1 に B1:B10 の合計を SUM 数式で入れて
You> ヘッダー行を太字にして背景を青にして
You> 「集計」という名前の新しいシートを作成して

You> report.docx を開いてドキュメントの構造を見せて
You> 「Q3の結果」を「Q3・Q4の結果」に置き換えて
You> インデックス3の段落の後に「以下に補足事項を記載する。」を挿入して
You> logo.png を文書の先頭に幅12cmで挿入して
You> 四半期売上データのテーブルを追加して
```

## 組み込みコマンド

| コマンド | 説明 |
|---------|------|
| `/reset` | 会話履歴をリセット |
| `/ls [パス]` | ファイル一覧を表示（Excel・Word ファイルをハイライト） |
| `/cd <パス>` | 作業ディレクトリを変更 |
| `/cwd` | 現在の作業ディレクトリを表示 |
| `/quit` | 終了 |

## 対応プロバイダーとデフォルトモデル

| プロバイダー | フラグ | デフォルトモデル | APIキー環境変数 |
|------------|--------|----------------|--------------|
| Anthropic | `--provider anthropic` | `claude-sonnet-4-6` | `ANTHROPIC_API_KEY` |
| Google Gemini | `--provider gemini` | `gemini-2.0-flash` | `GEMINI_API_KEY` |
| OpenRouter | `--provider openrouter` | `anthropic/claude-3.5-sonnet` | `OPENROUTER_API_KEY` |
| Ollama | `--provider ollama` | `qwen2.5:7b` | 不要 |

## Excel ツール一覧（13種類）

| ツール | 説明 |
|--------|------|
| `open_excel` | ファイルを開く・新規作成 |
| `list_sheets` | シート一覧を取得 |
| `read_sheet` | シート内容を読み取る（範囲指定可） |
| `read_cell` | 特定セルの値を読み取る |
| `write_cell` | セルに値を書き込む |
| `write_range` | 開始セルから 2D データを書き込む |
| `apply_formula` | セルに数式を設定する |
| `create_sheet` | 新しいシートを追加する |
| `delete_sheet` | シートを削除する |
| `format_cell` | 太字・色・背景・配置を設定する |
| `set_column_width` | 列幅を設定する |
| `save_excel` | 保存・別名保存 |
| `get_sheet_info` | 行数・列数・使用範囲を取得 |

## Word ツール一覧（13種類）

| ツール | 説明 |
|--------|------|
| `open_word` | ファイルを開く・新規作成 |
| `read_document` | 段落インデックス付きで全文を読み取る |
| `read_paragraph` | 特定段落の内容と書式を読み取る |
| `append_paragraph` | 末尾に段落を追加する |
| `insert_paragraph` | 指定インデックスの前後に段落を挿入する |
| `replace_text` | テキストを検索・置換する |
| `delete_paragraph` | 段落を削除する |
| `format_paragraph` | 段落の書式を変更する |
| `insert_image` | 画像を挿入する |
| `add_table` | 2D 配列からテーブルを追加する |
| `add_heading` | 見出しを追加する（H1〜H9） |
| `save_word` | 保存・別名保存 |
| `get_document_info` | 見出し一覧・段落数・テーブル数を取得 |

## プロジェクト構成

```
vibe-office/
├── main.py          # チャット UI・CLI エントリーポイント
├── agent.py         # Claude API / OpenAI 互換エージェントループ
├── excel_tools.py   # Excel 操作（openpyxl）
├── word_tools.py    # Word 操作（python-docx）
├── pyproject.toml   # プロジェクト設定（uv）
├── .env.example     # APIキー設定のテンプレート
└── .env             # APIキー（Git 管理外・各自で作成）
```
