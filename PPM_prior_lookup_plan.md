# PPM 前期ペアリング改善 実装計画 (案B)

作成日: 2026-03-31
対象ブランチ: edinet_api

---

## 背景・目的

現在のPPM成長率計算式:

```
成長率(N期) = value(N期, dim) / value(N-1期, dim) - 1
```

`value(N期)` と `value(N-1期)` は **別々の有報** から取得している。
セグメント組み替えがあった場合、N期のセグメント定義とN-1期のセグメント定義が
異なるため、成長率が正確でない。

同一有報内の前期データを使えば、セグメント一致が保証される:

```
2025年有報: [当期2025, 前期2024(組み替え後)]  ← 同一セグメント定義
→ 成長率 = value_2025_report(当期2025) / value_2025_report(前期2024)
```

---

## 実装方針 (案B: 分析シート構造を変えない)

### 概要

1. `convert_xbrl_to_excel.py` の ZIP処理ループ後（マージ後）に
   `ppm_prior_lookup` 辞書を構築し `segment_sheets_info` に追加
2. `segment_analysis.py` の PPM 関数に `ppm_prior_lookup` を追加引数として渡す
3. 成長率計算で `ppm_prior_lookup` が存在する年度はそちらを優先、
   存在しない場合は従来通りのフォールバック

分析シート（`_create_segment_analysis_sheet`）への変更なし。

---

## データ構造

### `ppm_prior_lookup` 辞書

```python
# キー: (segment_dim_label: str, current_period_date: str)
# 値:   {element_short_name: float}
#
# 例:
# ("ゲーム事業", "2025-03-31") -> {"NetSales": 120000.0, "OperatingIncome": 15000.0}
# ("ゲーム事業", "2024-03-31") -> {"NetSales": 100000.0, "OperatingIncome": 12000.0}
ppm_prior_lookup: dict[tuple[str, str], dict[str, float]]
```

キーの意味:
- `segment_dim_label`: セグメント名（例: "ゲーム事業"）
- `current_period_date`: **当期の期末日**（例: "2025-03-31"）
- 値の `float`: 同一有報の**前期**における当該セグメントの値

### なぜこの構造か

PPM計算時に「ある段の当期日付とセグメント名」から「同一有報の前期値」を引けばよい。
ネスト辞書にすることで、売上高・営業利益など複数の指標を一度に保持できる。

---

## 実装手順

### STEP 1: `ppm_prior_lookup` の構築
**ファイル**: `convert_xbrl_to_excel.py`
**場所**: マージループ（`for res in results:`）の後、`segment_sheets_info.append(...)` より前

#### 処理ロジック

```python
# --- PPM前期ペアリング辞書の構築 ---
# 各ZIPの results には res['values'] がある。
# res['values'][el][(std, dim, period)] = raw_value
# 1つのresから、セグメント軸のdimについて、
#   当期(最新のperiod)と前期(その次のperiod)のペアを特定する。

ppm_prior_lookup = {}  # {(dim, current_period): {element_short: float}}

for res in results:
    # このZIPの値を集める
    # セグメントdimごとに期末日のセットを集める
    seg_periods_by_dim = {}  # dim -> set of period dates
    for el, vals in res['values'].items():
        if el == '_metadata':
            continue
        for (std, dim, period), raw_val in vals.items():
            if not _is_segment_dim(dim):
                continue
            if dim not in seg_periods_by_dim:
                seg_periods_by_dim[dim] = set()
            seg_periods_by_dim[dim].add(period)

    # dimごとに当期(最新)と前期(その前)を特定
    for dim, periods in seg_periods_by_dim.items():
        sorted_periods = sorted(periods, reverse=True)
        if len(sorted_periods) < 2:
            continue
        current_period = sorted_periods[0]   # 当期 (最新)
        prior_period   = sorted_periods[1]   # 前期 (同一有報)

        # この(dim, current_period)の前期値を収集
        prior_vals = {}
        for el, vals in res['values'].items():
            if el == '_metadata':
                continue
            # elの短縮名(element short name)を取得
            el_short = el.split('_')[-1] if '_' in el else el
            for (std, d, period), raw_val in vals.items():
                if d == dim and period == prior_period:
                    try:
                        prior_vals[el_short] = float(raw_val)
                    except (TypeError, ValueError):
                        pass

        if prior_vals:
            key = (dim, current_period)
            if key not in ppm_prior_lookup:
                ppm_prior_lookup[key] = prior_vals
            # 既存エントリは新しいZIP(降順ソート済み)が優先 → 上書きしない
```

#### `segment_sheets_info` への追加

```python
segment_sheets_info.append({
    ...（既存のキー）...
    'ppm_prior_lookup': ppm_prior_lookup,   # ← 追加
})
```

---

### STEP 2: `add_segment_analysis_sheets` の変更
**ファイル**: `segment_analysis.py`
**場所**: `add_segment_analysis_sheets()` 内、PPM関数の呼び出し箇所

```python
# 変更前
_create_ppm_analysis_sheet(
    workbook=workbook,
    analysis_sheet_name=analysis_sheet_name,
    used_sheet_names=info['used_sheet_names'],
    debug_log=debug_log
)

# 変更後
_create_ppm_analysis_sheet(
    workbook=workbook,
    analysis_sheet_name=analysis_sheet_name,
    used_sheet_names=info['used_sheet_names'],
    ppm_prior_lookup=info.get('ppm_prior_lookup', {}),
    debug_log=debug_log
)
```

IFRS版 `_create_ppm_analysis_sheet_ifrs` も同様に変更。

---

### STEP 3: PPM関数のシグネチャ変更
**ファイル**: `segment_analysis.py`

```python
# 変更前
def _create_ppm_analysis_sheet(workbook, analysis_sheet_name, used_sheet_names, debug_log=None):

# 変更後
def _create_ppm_analysis_sheet(workbook, analysis_sheet_name, used_sheet_names,
                                ppm_prior_lookup=None, debug_log=None):
```

IFRSも同様:
```python
def _create_ppm_analysis_sheet_ifrs(workbook, analysis_sheet_name, used_sheet_names,
                                     ppm_prior_lookup=None, debug_log=None):
```

---

### STEP 4: 成長率計算での `ppm_prior_lookup` 使用
**ファイル**: `segment_analysis.py`
**場所**: 各PPM関数内の成長率配列構築箇所

#### 現在の成長率計算（概要）

分析シートの各列（セグメント）について:
- 当期売上 = 最新年度の値
- 前期売上 = その1つ前の年度の値
- 成長率 = 当期 / 前期 - 1

#### 変更後の計算

```python
# ppm_prior_lookupがある場合の優先ロジック
def _get_prior_value_with_lookup(dim_label, current_period, el_short, ppm_prior_lookup,
                                  fallback_value):
    """
    ppm_prior_lookupに当該(dim, current_period)のエントリがあれば優先使用。
    なければ fallback_value (分析シートのN-1期値) を返す。
    """
    if ppm_prior_lookup:
        key = (dim_label, current_period)
        if key in ppm_prior_lookup and el_short in ppm_prior_lookup[key]:
            return ppm_prior_lookup[key][el_short]
    return fallback_value
```

成長率の実際の計算箇所では:

```python
# dim_label: セグメント名（analysis_wsのヘッダー行から取得）
# current_period: 当期の期末日（analysis_wsのヘッダー行から取得）
# el_short: "NetSales" など

# 前期売上（従来の分析シートからの値）
prior_sales_fallback = _read_numeric(analysis_ws, sales_row_prior, col)

# ppm_prior_lookupを使った前期売上
prior_sales = _get_prior_value_with_lookup(
    dim_label, current_period, "NetSales",
    ppm_prior_lookup, prior_sales_fallback
)

# 成長率
growth_rate = current_sales / prior_sales - 1 if prior_sales else None
```

---

## 変更ファイルまとめ

| ファイル | 変更箇所 | 規模 |
|---|---|---|
| `convert_xbrl_to_excel.py` | マージループ後に `ppm_prior_lookup` 構築、`segment_sheets_info` に追加 | 中（40〜60行） |
| `segment_analysis.py` | PPM関数シグネチャ、`add_segment_analysis_sheets` の呼び出し、成長率計算ロジック | 中（50〜80行） |
| 他シート | 影響なし | なし |

---

## 実装時の注意点

### `res['values']` へのアクセスタイミング

`ppm_prior_lookup` の構築は、マージループが完全に終わった後では `results` リストはまだある（メモリ上）ため、再度ループ可能。ただし `_is_segment_dim` はマージループ内でクロージャとして定義されているので、同じ関数スコープ内で構築すること。

### セグメント名とelement shortの対応

`analysis_ws` のヘッダーはセグメント名（日本語ラベル）を持つ。
`ppm_prior_lookup` のキーもラベル名なので直接照合可能。

### フォールバック動作

`ppm_prior_lookup` が空、またはキーが存在しない場合は従来通りの成長率計算（分析シートのN-1期値）を使用。移行期・古いデータ・テストへの影響なし。

### IFRSと日本基準の両対応

`ppm_prior_lookup` は1回構築すれば両方の PPM 関数に渡せる（同一の `global_element_period_values` から構築するため）。

---

## 実装順序（推奨）

1. `convert_xbrl_to_excel.py`: `ppm_prior_lookup` 構築コードの追加（STEP 1）
2. `convert_xbrl_to_excel.py`: `segment_sheets_info.append` に `ppm_prior_lookup` キー追加
3. `segment_analysis.py`: `add_segment_analysis_sheets` の呼び出し変更（STEP 2）
4. `segment_analysis.py`: PPM関数シグネチャ変更（STEP 3）
5. `segment_analysis.py`: `_get_prior_value_with_lookup` ヘルパー追加・成長率計算変更（STEP 4）
6. 動作確認: セグメント組み替えがある企業（例: 光通信等）で前後比較

---

## 未解決課題

- `res['values']` の `raw_val` は文字列の場合もある（XBRLのfactは文字列で格納）。
  `float(raw_val)` 変換時に失敗するケースは `try/except` でスキップする。
- セグメント組み替えで**完全に新規**のセグメントが登場した場合、
  前期が存在しないため `ppm_prior_lookup` にもエントリなし → 成長率は計算不能（N/A）のまま。
  これは仕様通り（前期が存在しない以上、成長率は定義できない）。
- 同一期末日・同一セグメントのデータが複数ZIPに存在する場合、
  `results` は新しい年度順ソート済みのため、最初にヒットした（最新有報の）前期値を優先する。
