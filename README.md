# TDNet Scraper (`tdnet_scraper.py`)

TDNet（東証適時開示）の指定日付の **PDF（適時開示書類）** と **XBRL ZIP** を自動取得し、GitHub リポジトリへアップロードするスクレイパー。

---

## 特徴

| 機能 | 内容 |
|------|------|
| **JPX ticker フィルター** | JPX 上場銘柄リスト（`data_j.xls`）を参照し、国内株式のプライム・スタンダード・グロース銘柄だけを取得 |
| **フォルダ自動分轄** | 1 フォルダのファイルが `PART_SIZE`（デフォルト 1000）を超える場合、`yyyymmdd_Part1`/`yyyymmdd_Part2` … に自動分割 |
| **取り違え防止** | 表題リンク由来の ID のみ採用（行の混線で別ファイルを拾わない） |
| **英字入り ticker 対応** | `268A0 → 268A` / `33100 → 3310` など自動正規化 |
| **Preflight 書き込み確認** | 実行前に GitHub への書き込み可否を検証してから本処理を開始 |
| **magic bytes 検証** | ダウンロードしたファイルが本当に PDF / ZIP かをバイト列で確認 |
| **manifest.json** | 各フォルダに `manifest.json`（アップロード結果 + メタ情報）を自動生成 |
| **Colab / CLI 両対応** | Google Colab でも `python tdnet_scraper.py` でも動作 |

---

## リポジトリ構成

```
yukizi1113/tdnet
├── tekigikaizi/
│   ├── 20260226/              # ≦1000件の場合
│   │   ├── 1301_極洋_...pdf
│   │   └── manifest.json
│   ├── 20260207_Part1/        # 1001件以上の場合
│   │   ├── 9432_NTT_...pdf
│   │   └── manifest.json
│   └── 20260207_Part2/
│       └── manifest.json
├── XBRL/
│   ├── 20260226/
│   │   ├── 1301_極洋_...XBRL.zip
│   │   └── manifest.json
│   └── ...
├── __preflight__.txt          # 書き込みチェック用（自動生成）
└── tdnet_scraper.py
```

---

## インストール

```bash
pip install requests beautifulsoup4 lxml tqdm xlrd
```

- `xlrd`: JPX ticker フィルターに必要（未インストール時は対象外銘柄の混入を防ぐため停止）

---

## 使い方

### Google Colab（推奨）

1. `TARGET_DATE` を対象日に変更して、セルを実行するだけ。

```python
TARGET_DATE = "20260226"   # ← ここを変える
run(TARGET_DATE)
```

2. GitHub token は `GITHUB_TOKEN` 環境変数（Colab Secrets）、または実行時にペースト入力。

### CLI（コマンドライン）

```bash
# 環境変数にトークンをセット
export GH_TOKEN=ghp_xxxxxxxxxxxxxxxxxxxx

# 指定日を実行
python tdnet_scraper.py --date 20260226

# デバッグサンプルを無効化して静かに実行
python tdnet_scraper.py --date 20260226 --debug-sample 0

# フォルダ分轄のしきい値を変更（デフォルト1000）
python tdnet_scraper.py --date 20260226 --part-size 500
```

---

## CLI オプション

| オプション | デフォルト | 説明 |
|-----------|-----------|------|
| `--date` | `TARGET_DATE`（コード内の値） | 対象日 `yyyymmdd` |
| `--debug-sample` | `25` | 先頭 N 件をデバッグ表示（`0` で無効） |
| `--part-size` | `1000` | 1 フォルダあたりの最大ファイル数 |

---

## JPX ticker フィルター

実行時に JPX の [`data_j.xls`](https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls) をダウンロードし、D 列（市場・商品区分）が以下のいずれかに完全一致する銘柄だけを取得します。

| 取得対象区分 |
|---------------|
| プライム（内国株式） |
| スタンダード（内国株式） |
| グロース（内国株式） |

JPX ファイルの取得や解析に失敗した場合は、対象外銘柄をアップロードしないよう処理を停止します。

---

## フォルダ分轄の仕組み

```
PDF タスク数 ≦ PART_SIZE (1000)  →  tekigikaizi/20260207/
PDF タスク数 > PART_SIZE         →  tekigikaizi/20260207_Part1/
                                     tekigikaizi/20260207_Part2/
                                     ...

XBRL タスク数についても独立して同様の処理。
```

PDF と XBRL はそれぞれ独立して分轄されます。
各パートに `manifest.json` が生成されます。

---

## manifest.json の形式

```json
{
  "date": "20260226",
  "folder": "20260226",
  "root": "tekigikaizi/20260226/",
  "items": [
    {
      "ticker": "1301",
      "company": "極洋",
      "title": "2025年3月期 第3四半期決算短信",
      "time": "15:30",
      "file_id": "081220260226123456",
      "source_url": "https://www.release.tdnet.info/inbs/...",
      "github_path": "tekigikaizi/20260226/1301_極洋_....pdf",
      "status": "uploaded"
    }
  ]
}
```

`status` の値:

| 値 | 意味 |
|----|------|
| `uploaded` | 正常アップロード |
| `skipped_exists` | 既存ファイルあり（`OVERWRITE=False` 時） |
| `download_failed` | TDNet からのダウンロード失敗 |
| `type_mismatch` | magic bytes が期待した形式と異なる |
| `upload_failed` | GitHub API へのアップロード失敗 |

---

## 処理フロー

```
実行開始
  │
  ├─ [Step 1] build_target_tickers() JPX data_j.xls → 取得対象ticker集合
  │
  ├─ [Step 2] preflight_write_check()  GitHub書き込み権限確認
  │
  ├─ [Step 3] TDNet一覧ページを全ページ取得・パース
  │             I_list_001_yyyymmdd.html ～ (404になるまで)
  │
  ├─ [Step 4] JPX フィルター適用
  │             ticker が取得対象集合に含まれる行だけを選択
  │
  ├─ [Step 5] タスクリスト作成
  │             PDF タスク: (row, pdf_id) のリスト
  │             XBRL タスク: (row, xbrl_id) のリスト
  │
  ├─ [Step 6] Part 分轄の決定
  │             PART_SIZE を超える場合に _Part1/_Part2… を設定
  │
  ├─ [Step 7] 各チャンクをアップロード
  │             Download → magic bytes 検証 → GitHub PUT
  │
  └─ [Step 8] manifest.json 保存
```

---

## 設定値（コード内定数）

| 定数 | デフォルト | 説明 |
|------|-----------|------|
| `GITHUB_OWNER` | `"yukizi1113"` | GitHub オーナー |
| `GITHUB_REPO` | `"tdnet"` | GitHub リポジトリ名 |
| `TARGET_DATE` | `"20260226"` | デフォルト対象日（`--date` で上書き可） |
| `PDF_DIR_IN_REPO` | `"tekigikaizi"` | PDF の格納ディレクトリ |
| `ZIP_DIR_IN_REPO` | `"XBRL"` | XBRL ZIP の格納ディレクトリ |
| `OVERWRITE` | `True` | 同名ファイルの上書き |
| `SLEEP_SEC` | `1.0` | リクエスト間隔（秒） |
| `PART_SIZE` | `1000` | 1フォルダの最大ファイル数 |

---

## GitHub token の設定方法

| 方法 | 説明 |
|------|------|
| 環境変数 `GITHUB_TOKEN` | `export GITHUB_TOKEN=ghp_xxx` |
| 環境変数 `GH_TOKEN` | `export GH_TOKEN=ghp_xxx` |
| Colab Secrets | Colab の左パネル「🔑」に `GITHUB_TOKEN` を登録 |
| 対話入力 | 上記が未設定の場合、実行時にプロンプトが表示される |

必要なスコープ: **`repo`**（private repo の場合）または **`public_repo`**（public repo の場合）

---

## 注意事項

- TDNet のコンテンツは公開後一定期間（約 35 日程度）で削除される場合があります。毎営業日実行することを推奨します。
- GitHub Contents API の 100MB 制限のため、単一ファイルが 100MB を超える場合はスキップされます（実際には発生しないサイズです）。
- `SLEEP_SEC=1.0` は TDNet サーバーへの負荷軽減のためです。短縮する場合は自己責任で。
