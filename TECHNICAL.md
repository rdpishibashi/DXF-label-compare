# DXF-label-compare 実装仕様書（sonnet 引き継ぎ用）

このドキュメントは、`DXF-label-compare` アプリを **ゼロから実装する担当（sonnet モデル）**
向けの、自己完結した実装指示です。この 1 ファイルと、下記スキル・参照プロジェクトだけで
実装〜テスト〜Git 統合まで完了できるように書いてあります。

> **前提の役割分担**: Phase 1（事前調査・仕様確定）は完了済み。あなた（sonnet）は
> **Phase 2（実装）以降**を担当する。仕様はユーザー承認済みなので、仕様自体の
> 再確認は不要。実装を進めてよい。判断に迷う箇所のみユーザーへ確認する。

---

## 0. 最初に読むもの

1. `~/.claude/skills/dev-workflow/` （このスキルの Phase 2・Phase 3 手順に従う）
2. `~/.claude/skills/streamlit/` （UI 実装パターン。**特に §2 width, §3 アップロード,
   §4 Excel 出力, §6 セッション状態, §11 テーマ/ボタン, §12 動作確認, §13 落とし穴**）
3. `../CLAUDE.md`（Tools 共通ガイド。Excel 出力パターン・色分け規約）
4. 参照実装プロジェクト: `../DXF-extract-labels/`
   - `.streamlit/config.toml`, `.gitignore`, `requirements.txt`, `model/` 構成、
     `app.py` のボタン/セッション/ダウンロードの書き方を**踏襲**する

---

## 1. 目的とスコープ

`DXF-extract-labels` が出力した **2 つの Excel（A・B）の `Total` シートのラベルを比較**し、
差分（A のみ／B のみ／両方）を Excel として出力する Streamlit アプリ。

- 入力: Excel 2 個（A 用・B 用）
- 比較対象: 各 Excel の **`Total` シートの `ラベル` 列のみ**（`個数` は比較しないが出力する）
- 出力: 差分 Excel（差分シート＋サマリーシート）＋ 画面での色分け表示
- **単一アクション型ツール**（入力→実行→結果の 1 往復）。Step 番号見出しは付けない
  （streamlit スキル §11）。

---

## 2. 入力仕様（`Total` シートの実構造）

`DXF-extract-labels` の出力 Excel は `Summary` / `Total` / 図番ごとの個別シートを持つ。
本アプリが読むのは **`Total` シートのみ**。実測した構造:

| 列名 | 内容 | 本アプリでの扱い |
|------|------|-----------------|
| `ラベル` | ラベル文字列 | 比較キー（正規化後）・出力する |
| `個数` | 出現総数（int） | 比較しない・出力する（A個数/B個数） |
| `図番` | 出現図番のカンマ区切り | **読み込むが出力しない** |

- ヘッダーは 1 行目。データは 2 行目以降。
- **`Total` シート内でラベルは既にユニーク**（実測: 重複ゼロ・空欄ゼロ）。ただし
  正規化（§3）で別ラベルが同一化する可能性があるため、**正規化後キーで `個数` を
  合算**して安全にユニーク化すること。
- `Total` シートが存在しない Excel をアップロードした場合は、`ValueError` を送出し、
  `app.py` 側で `st.error("...には Total シートがありません")` を表示する
  （処理は止める）。列 `ラベル`・`個数` が無い場合も同様にエラー。

---

## 3. 正規化・比較ルール（重要）

### 3-1. ラベル正規化（比較前に必ず適用）

**全角 ASCII 文字があれば半角に変換**してから比較する（ユーザー指定:「日本語以外の
ラベルは、全角があれば半角にして比較」）。日本語（ひらがな・カタカナ・漢字）は変換しない。

- 変換対象: 全角 ASCII `U+FF01–U+FF5E` → 半角 `U+0021–U+007E`（`ord - 0xFEE0`）、
  全角スペース `U+3000` → 半角スペース `U+0020`
- 変換しない: ひらがな・カタカナ・漢字・その他の文字はそのまま
- 文字単位で適用する（混在ラベル `ＣＮ１番` → `CN1番` のように、ASCII 部分だけ半角化）

```python
def normalize_label(s: str) -> str:
    """全角ASCII(U+FF01-FF5E)を半角に、全角スペースを半角スペースに変換する。
    日本語文字（かな・カナ・漢字）は変換しない。"""
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        elif o == 0x3000:
            out.append(' ')
        else:
            out.append(ch)
    return ''.join(out)
```

> 注: `unicodedata.normalize('NFKC', s)` は使わない。NFKC は半角カナ→全角カナ変換や
> 合字分解など**日本語側にも副作用**があるため。上記の限定的な変換のみ行う。

### 3-2. 比較ルール

- ラベルの**完全一致**（正規化後の文字列が等しいかどうか）。
- `個数` は比較に使わない。「両方」に分類されたラベルで A個数 と B個数 が異なっていても、
  差分としてフラグは立てない（区分は「両方」のまま）。
- 表示・出力するラベル文字列は **正規化後（半角化後）の形**を使う（比較キーと一致させる）。

### 3-3. 区分

正規化後キーの集合演算で 3 分類:

| 区分ラベル | 条件 |
|-----------|------|
| `A のみ` | A に存在し B に無い |
| `B のみ` | B に存在し A に無い |
| `両方` | A・B 両方に存在 |

---

## 4. 出力仕様

### 4-1. 差分シート（シート名: `差分`）

- **1 枚に全ラベルを列挙**。`ラベル` で**昇順ソート**（Python 既定の文字列ソート＝
  コードポイント順。`DXF-extract-labels` と同じ方針）。
- 列（この順・この名前）:

  | ラベル | 区分 | A個数 | B個数 |
  |--------|------|-------|-------|

  - `A個数`/`B個数` は、その側に無いラベルでは**空欄**（Excel はセル未書き込み or `""`、
    DataFrame では `pd.NA`）。
- **行の色分け**（区分に応じて全 4 列の背景色を変える。ユーザー指定の配色）:

  | 区分 | 色 | 背景 hex | 文字 hex |
  |------|----|---------|---------|
  | `両方` | 緑 | `#C6EFCE` | `#006100` |
  | `A のみ` | 青 | `#D9E1F2` | `#1F4E79` |
  | `B のみ` | 茶色 | `#E2CFC0` | `#7F4F24` |

  （背景は薄色・文字は濃色で可読性を確保。この hex を既定とする。）
- ヘッダー: 太字・背景 `#4472C4`・白文字（Tools 共通の Excel ヘッダー書式）。
- `freeze_panes(1, 0)` で見出し行固定。列幅は内容に合わせて調整（`ラベル` は広め）。
- オートフィルタ（`autofilter`）を付けてよい。

### 4-2. サマリーシート（シート名: `サマリー`、差分シートより前に配置）

2 列（`項目`・`値`）で以下を出力:

| 項目 | 値（例: 339 vs 405 の実測ベースライン） |
|------|------|
| A ファイル名 | （アップロード時の元ファイル名） |
| B ファイル名 | （同上） |
| A ユニークラベル数 | 2666 |
| B ユニークラベル数 | 3533 |
| A のみ | 1847 |
| B のみ | 2714 |
| 両方 | 819 |
| ユニーク合計 | 5380 |

### 4-3. 画面表示

- 差分テーブルを `st.dataframe` で表示。**§4-1 と同じ色分け**を `pandas.Styler`
  （`background-color` / `color` を返す関数）で適用。`hide_index=True`、`width='stretch'`。
- サマリー件数は `st.info` などで簡潔に併記してよい。
- Excel ダウンロードは `st.download_button`（`type="primary"`, `width='stretch'`,
  mime は xlsx）。既定ファイル名 `label_compare.xlsx`。

---

## 5. ファイル構成（3 層構造）

`DXF-extract-labels` に倣い、モデル層（比較ロジック）を Streamlit 非依存の純関数で切り出す。

> 以下は初回実装（§1〜§11）時点の構成。その後 §12・§13 の追加機能で
> `model/drawing_filter.py`・`model/same_workbook.py`・`model/same_workbook_output.py`
> と対応するテストが増えている。現在の実際の構成は `ls model/ tests/unit/` で確認すること。

```
DXF-label-compare/
├── app.py                    # View層: Streamlit UI（薄く保つ）
├── requirements.txt
├── .gitignore                # ../DXF-extract-labels/.gitignore をコピー
├── .streamlit/
│   └── config.toml           # ../DXF-extract-labels/.streamlit/config.toml と同一
├── model/
│   ├── __init__.py           # 空ファイル
│   ├── excel_input.py        # Model層: Total シート読み込み → 正規化 dict
│   ├── compare_labels.py     # Model層: 正規化・差分・サマリー（純関数・テスト対象）
│   └── excel_output.py       # Model層: 差分 Excel 生成
├── tests/
│   └── unit/
│       └── test_compare_labels.py
├── TECHNICAL.md               # 本ファイル（旧 docs/IMPLEMENTATION_SPEC.md、2026-07-16 に
│                               # 他プロジェクトと同じ命名規則でプロジェクトルートへリネーム）
└── README.md                 # 日本語ドキュメント
```

`requirements.txt`（ezdxf は不要。入力は Excel）:

```
streamlit>=1.40.0
pandas>=2.0.0
openpyxl>=3.1.0
xlsxwriter>=3.0.0
```

`.streamlit/config.toml` と `.gitignore` は参照プロジェクトからコピー:

```bash
cp ../DXF-extract-labels/.streamlit/config.toml .streamlit/config.toml
cp ../DXF-extract-labels/.gitignore .gitignore
```

---

## 6. 各モジュールの責務と関数シグネチャ（参考実装つき）

### 6-1. `model/compare_labels.py`（純関数・単体テストの主対象）

```python
import pandas as pd

def normalize_label(s: str) -> str:
    ...  # §3-1 の実装をそのまま

def compare_labels(a: dict[str, int], b: dict[str, int]) -> pd.DataFrame:
    """正規化済みの {ラベル: 個数} 2 つを比較し、差分 DataFrame を返す。
    - columns: ['ラベル', '区分', 'A個数', 'B個数']（この順）
    - 区分 ∈ {'A のみ', 'B のみ', '両方'}
    - 無い側の個数は pd.NA
    - ラベル昇順（sorted）
    """
    labels = sorted(set(a) | set(b))
    rows = []
    for lbl in labels:
        in_a, in_b = lbl in a, lbl in b
        kubun = '両方' if (in_a and in_b) else ('A のみ' if in_a else 'B のみ')
        rows.append({
            'ラベル': lbl,
            '区分': kubun,
            'A個数': a.get(lbl, pd.NA),
            'B個数': b.get(lbl, pd.NA),
        })
    return pd.DataFrame(rows, columns=['ラベル', '区分', 'A個数', 'B個数'])

def summarize(df: pd.DataFrame, a_name: str, b_name: str) -> dict:
    """サマリー用の件数集計を返す。"""
    a_only = int((df['区分'] == 'A のみ').sum())
    b_only = int((df['区分'] == 'B のみ').sum())
    both   = int((df['区分'] == '両方').sum())
    return {
        'A ファイル名': a_name,
        'B ファイル名': b_name,
        'A ユニークラベル数': both + a_only,
        'B ユニークラベル数': both + b_only,
        'A のみ': a_only,
        'B のみ': b_only,
        '両方': both,
        'ユニーク合計': len(df),
    }
```

### 6-2. `model/excel_input.py`

```python
import io
import pandas as pd
from model.compare_labels import normalize_label

REQUIRED_SHEET = 'Total'

def load_total_labels(file_bytes: bytes) -> dict[str, int]:
    """Excel の Total シートを読み、正規化ラベル → 合計個数 の dict を返す。
    Total シートが無い / 必要な列が無い場合は ValueError。"""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if REQUIRED_SHEET not in xls.sheet_names:
        raise ValueError(f"'{REQUIRED_SHEET}' シートが見つかりません")
    df = xls.parse(REQUIRED_SHEET)
    if 'ラベル' not in df.columns or '個数' not in df.columns:
        raise ValueError("Total シートに『ラベル』『個数』列がありません")
    agg: dict[str, int] = {}
    for lbl, cnt in zip(df['ラベル'], df['個数']):
        if pd.isna(lbl):
            continue
        key = normalize_label(str(lbl))
        agg[key] = agg.get(key, 0) + (int(cnt) if pd.notna(cnt) else 0)
    return agg
```
> `st.cache_data` を使う場合は `file_bytes: bytes` を引数にする（streamlit スキル §5。
> アンダースコア引数は使わない）。

### 6-3. `model/excel_output.py`

- `xlsxwriter` エンジンで `サマリー`→`差分` の順に書く。
- 区分ごとの `workbook.add_format({'bg_color': ..., 'font_color': ...})` を 3 つ用意し、
  各データ行を区分に応じた書式で 4 列とも書き込む（§4-1 の hex）。
- ヘッダー書式・列幅・`freeze_panes(1,0)`・`autofilter`。
- 戻り値は `bytes`（`io.BytesIO().getvalue()`）。streamlit スキル §4 の
  `create_excel` パターンに従う。
- **空欄セル**（無い側の個数）は書き込まない or `''` を書く（数値 0 を書かないこと。
  「0 個」と「存在しない」を区別するため空欄にする）。

### 6-4. `app.py`（View 層・薄く）

- `st.set_page_config(page_title="DXF Label Compare", page_icon="🔍", layout="wide")`
- 2 つの `st.file_uploader`（`type=['xlsx']`, 各 1 ファイル）。ラベル例:
  「A: 比較元の Excel（extract-labels 出力）」「B: 比較先の Excel」
- `比較` ボタン: `type="primary"`, `disabled=（A・B 両方が揃っていないと無効）`。
- 実行時: `load_total_labels` を A・B に適用 → `compare_labels` → 結果を
  `st.session_state` に保存（`UploadedFile` は 1 rerun で失効するため、
  `uploaded.getvalue()` で bytes 化してから渡す。streamlit スキル §3・§6）。
- 結果表示: Styler で色分けした `st.dataframe` ＋ サマリー ＋ ダウンロードボタン。
- 元ファイル名は `uploaded.name` で保持し、サマリー・DL に使う。
- エラーは `try/except ValueError` で `st.error` 表示。

Styler の色分け（画面表示用）:

```python
_BG = {'両方': '#C6EFCE', 'A のみ': '#D9E1F2', 'B のみ': '#E2CFC0'}
_FG = {'両方': '#006100', 'A のみ': '#1F4E79', 'B のみ': '#7F4F24'}

def _row_style(row):
    k = row['区分']
    css = f'background-color: {_BG[k]}; color: {_FG[k]}'
    return [css] * len(row)

styled = df.style.apply(_row_style, axis=1)
st.dataframe(styled, width='stretch', hide_index=True)
```

---

## 7. テスト

### 7-1. 単体テスト `tests/unit/test_compare_labels.py`

最低限、以下を検証:

1. **基本の区分**:
   ```python
   a = {'CN1': 2, 'R10': 1, 'ABC': 3}
   b = {'CN1': 5, 'X1': 1, 'ABC': 4}
   df = compare_labels(a, b)
   # A のみ: R10 / B のみ: X1 / 両方: ABC, CN1
   # ソート順: ABC, CN1, R10, X1
   ```
   - 区分の割り当て、A個数/B個数 の値、無い側が pd.NA、ソート順を assert。
2. **個数が違っても両方**: `a={'ABC':1}`, `b={'ABC':99}` → 区分 `両方`。
3. **正規化（全角→半角）**: `normalize_label('ＣＮ１') == 'CN1'`、
   `normalize_label('AB１２') == 'AB12'`、`normalize_label('あ　い') == 'あ い'`
   （全角スペース→半角、かなは不変）、`normalize_label('抵抗Ｒ') == '抵抗R'`。
4. **正規化による同一化と個数合算**: `excel_input` レベル、または dict 構築時に
   `ＣＮ１`（全角）と `CN1`（半角）が同一キーに合算されること。
5. **`summarize`** の件数が整合すること。

`pytest tests/unit/ -v` で実行。

### 7-2. ブラックボックステスト（実データ・期待値ベースライン）

ユーザー提供の実ファイル 2 つを A・B として使う:

- A = `/Users/ryozo/Downloads/extracted_labels-339_Unit内結線図.xlsx`
- B = `/Users/ryozo/Downloads/extracted_labels-405_展開接続図.xlsx`

**期待される集計値（この 2 ファイルで検証済み・正規化適用後）**:

| 項目 | 値 |
|------|-----|
| A ユニークラベル数 | 2666 |
| B ユニークラベル数 | 3533 |
| A のみ | 1847 |
| B のみ | 2714 |
| 両方 | 819 |
| ユニーク合計（差分シート行数） | 5380 |

- まずモデル層だけで（Streamlit を起動せず）上記の値が再現することを確認する。
- 次に streamlit スキル §12 の手順で実アプリを起動し（`chromium-cli` が無ければ
  Playwright + `channel="chrome"` のフォールバック）、両ファイルをアップロード→
  「比較」→ ダウンロードした xlsx を `openpyxl` で開き、**シート構成・行数(5380)・
  区分ごとの件数・色書式**まで検証する。スクリーンショットの目視だけで終わらせない。
- 注: この 2 ファイルには全角 ASCII ラベルは存在しないため、正規化の有無で件数は
  変わらない（正規化ロジックの検証は §7-1 の単体テストで担保する）。

回帰テストとして残す場合は `tests/regression/` に保存し、実行前にユーザーへ確認する
（dev-workflow Phase 2-4）。

---

## 8. 実装の組み合わせ表（dev-workflow Phase 2-1・必須）

今回の変更が触れるコードパスの直積。すべて新規実装なので「影響あり＝要テスト」。

| 軸: 入力状態 \ 区分 | A のみ | B のみ | 両方 |
|---|---|---|---|
| 通常ラベル | 要テスト(§7-1 #1) | 要テスト(§7-1 #1) | 要テスト(§7-1 #1) |
| 個数が異なる | ― | ― | 要テスト(§7-1 #2) |
| 全角ASCII含む | 要テスト(§7-1 #3,#4) | 要テスト(§7-1 #3,#4) | 要テスト(§7-1 #3,#4) |
| 実データ全体 | 1847(§7-2) | 2714(§7-2) | 819(§7-2) |

異常系: `Total` 無し / 列不足 → `ValueError` → `st.error`（app 側で確認）。

---

## 9. Git ワークフロー（dev-workflow Phase 1-5・Phase 3）

このフォルダはまだ git 未初期化（Tools 配下の各サブフォルダは独立 repo）。

1. **Phase 1-5**: `git init` → 初期コミット（空 or この spec のみ）→
   作業ブランチ `git checkout -b feature/initial-implementation` →
   `git tag baseline-YYYYMMDD`。リモート未設定なので fetch/pull はスキップ。
2. **Phase 2**: 小さいステップでコミット（動作確認済みのみ）。モデル層→出力→UI の順。
3. **Phase 3-2（必須ゲート・main マージ前）**: `README.md` を日本語で作成し、
   `docs/` を整備。`ls *.md docs/*.md` で対象を列挙し、各ファイルを更新してから 3-3 へ。
   - `README.md`: 目的・使い方・入出力・Total シート前提・正規化ルール・配色を記載。
   - この実装仕様書は実装後 `TECHNICAL.md`（プロジェクトルート、Tools 配下の他プロジェクトと
     同じ命名規則）へ発展させるか、README に要点を吸収したうえで残置してよい。
4. **Phase 3-3**: `main` へ `--no-ff` マージ。
5. **Phase 3-5**: push はユーザー確認後（リモート未設定ならセットアップをユーザーに案内）。

> 参照元 `../CLAUDE.md` の「共有 `model/extract_labels.py`」の伝播ルールは
> **本アプリには無関係**（extract_labels を使わない）。新規の独立コードなので
> 他プロジェクトへの伝播は不要。

---

## 10. 完成チェックリスト

- [x] `.streamlit/config.toml` / `.gitignore` をコピー、`requirements.txt` 作成
- [x] `model/compare_labels.py`（normalize_label / compare_labels / summarize）
- [x] `model/excel_input.py`（load_total_labels、Total 欠如で ValueError）
- [x] `model/excel_output.py`（サマリー＋差分、区分別 3 色、空欄セル、freeze/filter）
- [x] `app.py`（2 uploader・比較ボタン primary+disabled・Styler 色分け表示・DL primary）
- [x] 単体テスト（§7-1 の 1〜5）が pass
- [x] 実データで A のみ1847 / B のみ2714 / 両方819 / 合計5380 を再現
- [x] 実アプリ起動 → DL xlsx を openpyxl で開き行数・色書式まで検証
- [x] README.md 作成（Phase 3-2 ゲート）→ feature ブランチにコミット
- [ ] main へ `--no-ff` マージ、中間生成物削除、push はユーザー確認後

## 11. 実装時に判明した追加事項（sonnet 実装ログ）

- **画面表示での `pd.NA` が `"None"` と表示される問題**: §6-4 で示した
  `Styler.format(na_rep=...)` は、`st.dataframe`（Streamlit 1.54 の Arrow ベース
  データグリッド）には効かないことが実機検証で判明した（`column_config.NumberColumn`
  でも同様に抑制不可）。`A個数`/`B個数` を `pandas.Int64`（nullable整数型）にした上で、
  **画面表示専用**に文字列化したコピー（欠損は空文字）を作り、それを `st.dataframe` に
  渡す方式で解決した（`app.py` の `_for_display()`）。Excel 出力（`write_blank` 使用）
  は最初から正しく空欄だったため影響なし。詳細は README.md「既知の制約」を参照。
  - トレードオフ: 表示用コピーは文字列列になるため、数値の右寄せ（CSS
    `text-align: right` は `st.dataframe` の Styler 経由では反映されない）は効かず
    左寄せになる。spec に右寄せの明示要件は無かったため許容した。

## 12. 追加仕様（2026-07-12・初回リリース後の追加要望）

初回実装（§1〜§11、main へマージ済み）の後、ユーザーから以下の追加要望を受けて実装した。

1. **A/B の表示名を固定**: A のアップローダーラベルを「展開接続図」、B を
   「UNIT内結線図」に変更（`app.py` の `file_uploader` ラベルのみ変更。内部の
   区分値 `A のみ`/`B のみ`/`両方` や列名 `A個数`/`B個数` は変更していない）。
2. **B側の図番フィルタ**: B の `Summary` シートの `タイトル` 列を使い、
   「UNIT内結線図のみ」（既定）／「UNIT内結線図以外」／「全部」の3択で比較対象の
   図番を絞り込めるようにした。詳細仕様は README.md「Bの図番フィルタ」を参照。

追加した/変更したモジュール:

- `model/drawing_filter.py`（新規）: `select_drawing_numbers()`（タイトル→対象図番集合）、
  `aggregate_filtered_rows()`（図番リストでの絞り込み集計）。純関数、単体テスト
  `tests/unit/test_drawing_filter.py` あり。
- `model/excel_input.py`（拡張）: `load_total_rows()`（Total シートを図番付きで
  行単位に返す）、`load_summary_titles()`（Summary シートから 図番→タイトル の dict）
  を追加。既存の `load_total_labels()`（A側で使用、図番なし集約）は変更なし。
- `model/compare_labels.py`: `summarize()` に任意引数 `b_filter_mode` を追加
  （指定時のみ `B ファイル名` の直後に `B 絞り込み条件` を挿入。デフォルト `None` で
  後方互換）。
- `app.py`: B側アップローダーの下に `st.radio` で絞り込み選択肢を表示（既定値
  「UNIT内結線図のみ」）。実行時に A は従来どおり `load_total_labels`、B は
  `load_summary_titles` + `load_total_rows` + `select_drawing_numbers` +
  `aggregate_filtered_rows` の組み合わせで絞り込み後の dict を作る。

設計判断の記録:

- **個数は絞り込み後も再計算しない**: 選択した図番だけに絞ってラベルを対象化するが、
  そのラベルの `個数` は `Total` シートに記録された値（全図番分の合計）をそのまま使う。
  ユーザー指示が「Total シートの図番欄」を絞り込みキーとして明示していたため、
  個別図面シートを読んで再集計するような追加実装はスコープ外と判断した（個数は
  そもそも比較に使わない参考値のため、この簡略化は既存方針と整合する）。
- **一致判定の正規化に既存の `normalize_label()` を再利用**: 「ＵＮＩＴ内結線図」
  （全角）と「UNIT内結線図」（半角）を同一視する要件は、ラベル比較で既に使っている
  全角ASCII→半角の正規化ルールとまったく同じ性質だったため、新しい正規化関数を
  作らず `compare_labels.normalize_label()` をそのまま流用した。実データ（B側サンプル、
  339ファイル）で実際に両表記が混在していることを確認済み（全角13件・半角3件）。
- **実データでの検証値**（A=405_展開接続図.xlsx、B=339_Unit内結線図.xlsx、
  全71図番中UNIT内結線図タイトルは16図番）:

  | モード | A のみ | B のみ | 両方 | B ユニークラベル数 |
  |--------|-------|-------|------|-------------------|
  | UNIT内結線図のみ | 2769 | 409 | 764 | 1173 |
  | UNIT内結線図以外 | 3313 | 1471 | 220 | 1691 |
  | 全部 | 2714 | 1847 | 819 | 2666 |

  「全部」モードの値は、初回実装時の基準値（旧A=339/旧B=405、A/B入れ替え前:
  A のみ=1847・B のみ=2714・両方=819・合計=5380）と対応関係が一致することを
  確認済み。旧基準は339をA・405をBとしていたのに対し、本セッションでは
  405をA・339をBとしている（A/Bが入れ替わっている）ため、旧「A のみ=1847」
  （339固有ラベル数）は新「B のみ=1847」に、旧「B のみ=2714」（405固有ラベル数）
  は新「A のみ=2714」に、それぞれ対応する。両方=819・合計=5380は入れ替えの
  影響を受けないため両セッションで同一。

## 13. タブ名変更・streamlit スキル準拠スタイル適用（2026-07-16）

ユーザーから以下の要望を受けて `app.py`（View層のみ）を変更した。Model層（`model/`）は無変更。

1. **タブ名変更・表示順変更**: 「同じExcel内で比較」→「結線図-組立図比較」、
   「2つのExcelを比較」→「展開図-結線図比較」にリネームし、表示順を
   展開図-結線図比較 → 結線図-組立図比較 に変更（旧: 結線図-組立図比較が先頭だった）。
   各タブ内の説明文・アップローダーラベル（A: 展開接続図のラベル・ファイル / B: UNIT内
   結線図セットのラベル・ファイル）・実行ボタン文言もタブ名に統一。
   内部識別子（`session_state` キー・関数名・ダウンロードファイル名・列名 `A個数`/`B個数`
   等）は変更していない。
2. **タイトル変更**: `st.title("DXF Label Compare")` →
   `st.title("DXF Label Compare - 機器符号比較")`。`st.set_page_config` の
   `page_title`（ブラウザタブ名）は変更対象外だったため `"DXF Label Compare"` のまま。
3. **streamlit スキル §11 準拠スタイルの適用**:
   - 箱型タブ CSS（`.stTabs` の `data-baseweb="tab"` セレクタでボーダー・角丸・
     選択状態の背景を制御）を `main()` 冒頭で `st.markdown(..., unsafe_allow_html=True)`
     により注入。
   - 日英混在フォントサイズ調整（`@font-face` + `unicode-range` + `size-adjust: 94%`
     による日本語グリフのみの縮小、フォント名 `AppMixedFont`）も同様に注入。
     初回実装時は未適用だったため、レビュー指摘を受けて追加した。
   - どちらもテーマ色ではなく無彩色 `rgba()` ベースのため、ライト/ダークテーマ双方で
     利用可能（streamlit スキルの指針どおり）。旧§12「見た目・タイポグラフィのCSS
     カスタマイズ集」は、この指摘を受けて streamlit スキル側で §11 に統合され、
     新規プロジェクトへのデフォルト適用対象へ格上げされた（詳細は §14 参照）。

ファイル名変更（ドキュメントのみ、コード非関連）:

- `docs/IMPLEMENTATION_SPEC.md` → `TECHNICAL.md`（プロジェクトルート）にリネーム。
  `Tools/` 配下の他プロジェクト（`DXF-extract-labels`・`DXF-diff-manager` 等）が
  技術文書をプロジェクトルート直下の `TECHNICAL.md` に置く命名規則と揃えるため。
  `docs/` フォルダは中身が無くなったため実質消滅（git は空ディレクトリを追跡しない）。

## 14. `utils/` → `model/` リネーム、および関連スキルの是正（2026-07-16）

§13 の作業後、ユーザーから「日英混合フォント対応と `model/` フォルダ名が毎回抜けがちなのは
なぜか」という指摘を受けて調査した結果、2つの独立した原因が判明した。

1. **日英混合フォント（streamlit スキル旧§12）**: スキル側が「§11は新規プロジェクトの
   既定スタイル」「§12は必要なプロジェクトだけ個別に適用する」と明記して区別しており、
   §12 は明示的な指示がない限り適用対象外という opt-in 表現になっていた。→ streamlit
   スキルの §12 を §11 に統合し、新規プロジェクトへの一律デフォルト適用に変更した
   （`~/.claude/skills/streamlit/SKILL.md`）。
2. **`model/` フォルダ名**: `Tools/CLAUDE.md` の記述は DXF-extract-labels・
   DXF-diff-manager の2プロジェクトが `utils/`→`model/` にリネームした**過去の事実**の
   記録に過ぎず、「今後の新規プロジェクトは `model/` にすべき」という指示にはなって
   いなかった。実際に新規プロジェクト作成時に参照する `dxf-new-project-scaffolding`
   スキル（`Tools/.claude/skills/dxf-new-project-scaffolding/SKILL.md`）は、
   今も `utils/__init__.py`・`utils/extract_labels.py` という旧命名のままだった。
   → scaffolding スキルを `model/` 命名に更新し、DXF-label-compare 自体も
   `utils/` → `model/` にリネームした（`git mv utils model`、`app.py`・`model/*.py`・
   `tests/unit/*.py` 内の `from utils.` → `from model.` を一括置換、`README.md`・
   `TECHNICAL.md` 内のパス表記も追随）。単体テスト17件 pass、実アプリ起動確認済み。

この節の教訓: 「複数プロジェクトが同じ変更をした」という**事実の記録**と、「今後の
新規プロジェクトはこうすべき」という**指示**は別物であり、前者を CLAUDE.md に書いただけ
では後者として機能しない。新しい規約を定着させるには、実際に参照される手順書
（ここでは scaffolding スキル・streamlit スキル）側を直接更新する必要がある。

## 15. 「結線図-組立図比較」の出力フォーマットを「展開図-結線図比較」に統一（2026-07-16）

ユーザーから「結線図-組立図比較の出力ファイルのフォーマット・シート構成が
展開図-結線図比較と異なる（Summaryシートがないなど）」という指摘を受けて対応した。

**変更前の差異**:
- サマリー情報が独立シートでなく「比較結果」シート内に結合セルのタイトル行・
  集計表として埋め込まれていた。
- 配色・ヘッダー書式が異なり（ヘッダー背景 `#5B9BD5`、区分ラベルも「共通/Aのみ/Bのみ」
  で展開図-結線図比較の「両方/A のみ/B のみ」と不一致）、グリッド線非表示・
  結合セルなど見た目のスタイルも別物だった。
- `A_ラベル一覧`・`B_ラベル一覧`・`A_対象図番` の3枚の補助シートを追加で持っていた。

**変更後**（ユーザー確認済みの方針）:
1. 補助3シートは削除し、`サマリー`＋`差分`の2シートのみに統一。
2. 区分値を「共通/Aのみ/Bのみ」→「両方/A のみ/B のみ」に統一（`model/same_workbook.py`）。
   これに伴い画面表示・`summary` 辞書のキー（`共通ラベル数`→`両方`）も追随。
3. `model/same_workbook_output.py` を `model/excel_output.py` と同じヘッダー書式
   （太字・背景 `#4472C4`・白文字・罫線）、同じ区分別配色（両方=緑`#C6EFCE`、
   A のみ=青`#D9E1F2`、B のみ=茶`#E2CFC0`）、同じ `freeze_panes(1,0)`・`autofilter`
   方針で全面書き直し。差分シートの列は「ラベル・A合計出現数・A図番数・B出現数・
   A図番・比較結果」の6列のまま維持（この比較固有の情報である A図番数・A図番は
   展開図-結線図比較の4列に合わせて削ることはしない、とユーザー確認済み）。
4. `model/same_workbook.py`: 削除した3シート専用に構築していた `a_rows`/`b_rows`/
   `selected_rows`（および使われなくなった `Drawing` データクラスの `title`/`subtitle`
   フィールド）を削除。`len()` は `a_labels`/`b_labels`/`selected_drawings` から
   直接取得するよう簡略化。`selected_drawings` も `Drawing` オブジェクトのリストから
   図番文字列のリストに単純化（`.title`/`.subtitle` を読む箇所が無くなったため）。

**両ファイルの見た目が再び乖離しないよう**、`model/same_workbook_output.py` の
モジュールdocstringに「`model/excel_output.py` と書式を揃えている。変更時は両方
同時に見直すこと」という運用上の注意を明記した。

単体テスト17件 pass（`tests/unit/test_same_workbook.py` の区分値・サマリーキーの
assertを更新、削除フィールドへのassertを削除）。実アプリで実際にExcelをダウンロードし、
`openpyxl` でシート構成（`サマリー`・`差分`のみ）・ヘッダー色（`#4472C4`）・
区分別の行色（緑/青/茶）を検証済み。

## 16. 「展開図-結線図比較」に「指定領域での比較」を追加（2026-07-23）

ユーザー要望: A・B両方が `DXF-extract-labels` の「領域を検出」オプションで出力した
Excel（`領域別ラベル一覧` シートを持つ）である場合、共通の領域名を選んでその領域内
だけをラベル比較できるようにしたい。

### 設計方針

既存4モジュールを拡張する形で実装した（新規モジュールは作らない、3層構造は不変）。

- **`model/excel_input.py`**: `has_region_sheet()`（シート名一覧のみ読む軽量存在
  チェック）、`load_region_rows()`（`領域別ラベル一覧`を
  `(正規化領域名, 正規化ラベル, 合計個数, 図番リスト)` のタプル列で返す）を追加。
  `領域別ラベル一覧` は列名 `図番`/`個数` がファイル数分繰り返される構成
  （pandas が `図番.1`/`個数.1` 等に自動リネームする）ため、列名ではなく
  **位置**（3列目以降を2列ずつ）で読む。図番リストには個数が1以上だった
  列の図番のみを含める（`load_total_rows` の図番リストと同じ「実際に出現した
  ファイルのみ」という考え方）。
- **`model/drawing_filter.py`**: `aggregate_region_rows()`を追加。
  `aggregate_filtered_rows()`と同じ絞り込み方針（selected_gzuban との重なりで
  行単位に採否を決め、**個数は絞り込み後も再計算しない**）を領域名ごとに適用する。
- **`model/compare_labels.py`**: `REGION_DIFF_COLUMNS`（`['領域名'] + DIFF_COLUMNS`）、
  `summarize_metrics()`（`summarize()`からファイル名部分を除いた集計だけを切り出す
  リファクタ。`summarize()`は内部でこれを呼ぶよう変更、公開契約は不変）、
  `compare_labels_by_region()`（領域名ごとに`compare_labels()`を実行し`領域名`列
  付きで結合、`(diff_df, metrics_by_region)`のタプルを返す）、
  `build_region_summary_rows()`（`項目`/`領域名`/`値`のサマリー行リストを構築）を追加。
- **`model/excel_output.py`**: `_write_diff_sheet`/`_write_summary_sheet`を
  列リスト・行リスト受け取りに一般化（既存の4列/2列シート出力の挙動は不変。
  内部を`_create_excel_output()`共通関数に統合）、
  `create_region_compare_excel_output()`を新設。
- **`app.py`**: `_show_region_selection()`を新設（A・B両方に`領域別ラベル一覧`が
  ある場合のみ「指定領域で比較する」チェックボックスを表示し、ONなら共通領域名の
  `st.multiselect`を表示）。「比較」ボタンのハンドラは`region_mode`で分岐し、
  `two_workbook_is_region_mode`をsession_stateに保存して結果表示を切り替える
  （DXF-extract-labelsの`is_region_mode`と同じ命名パターン）。

### 決定事項（ユーザー確認済み）

- **比較範囲**: ONのとき選択した領域名だけに限定する（`Total`シート全体の比較は
  行わない）。1ラベルが複数の選択領域に属す場合、領域ごとに独立した行として
  複数回出力される（`region_detector`の「1ラベルが複数領域に所属可」という
  設計と整合）。
- **B図番フィルタ**: 指定領域での比較にも同様に適用する。ただし共通領域名の
  一覧自体はフィルタと独立に（構造的な情報として）算出し、フィルタは比較実行時
  の個数集計にのみ適用する。
- **サマリー形式**: `項目`/`領域名`/`値`の3列構成にする（項目名に領域名を
  埋め込む案は採らなかった）。

### テスト

単体テスト16件追加（`tests/unit/test_excel_input.py`新設・
`tests/unit/test_drawing_filter.py`/`test_compare_labels.py`に追加・
`tests/unit/test_excel_output.py`新設）、既存17件と合わせて35件全てpass。
実データ（`DXF-extract-labels`の`sample-dxf/problems/EE6888-637-01A.dxf`・
`EE6491-039-21A.dxf`から生成した領域付きExcel、共通領域名`SYSTEM I/F BOX`）で
実アプリを起動し、チェックボックス表示→領域選択→比較→ダウンロードした
Excelを`openpyxl`で検証（差分シート5列・サマリーシート3列・領域ごとの
行ブロック・B絞り込み条件の反映）。正常系（領域モードOFF、`Total`シート
必須のエラー表示）も回帰確認済み。

### 副産物: DXF-extract-labels側の欠落を2件発見・修正

この機能の実データ検証中に、`DXF-extract-labels`側の以下の欠落を発見し、
別途修正した（詳細は同プロジェクトの`docs/VERSION_HISTORY.md` v1.9.9・
v1.9.10を参照）。

1. **region モードのSummaryシートに`図番`/`タイトル`列が無かった**
   （v1.9.9）: B図番フィルタが領域モード出力に対して機能しない原因だった。
2. **図面枠が見つからない図面のラベルが「機器符号（候補）以外も抽出」ON時に
   丸ごと消えていた**（v1.9.10）: `analyze_dxf_regions()`が図面枠エラー時に
   `labels`を空のまま返していたことが原因。本プロジェクトの機能とは直接
   関係ないが、実データ検証の過程で発覚しユーザー報告により対応した。

---

*作成: 2026-07-12 / Phase 1 担当（Opus）→ Phase 2 以降担当（sonnet）への引き継ぎ*
*実装完了: 2026-07-12 / Phase 2〜Phase 3-2 まで sonnet が実施*
*2026-07-16: タブ名変更・streamlit スキル準拠スタイル適用・`TECHNICAL.md` へのリネーム*
*2026-07-16: `utils/` → `model/` リネーム、streamlitスキル §12→§11統合、
dxf-new-project-scaffolding スキルの `model/` 命名反映*
*2026-07-16: 結線図-組立図比較の出力フォーマットを展開図-結線図比較に統一
（サマリー＋差分の2シート化、区分値「両方/A のみ/B のみ」への統一、配色統一）*
*2026-07-23: 「展開図-結線図比較」に「指定領域での比較」を追加（§16）。
既存4モジュールの拡張のみ、新規モジュールなし*
