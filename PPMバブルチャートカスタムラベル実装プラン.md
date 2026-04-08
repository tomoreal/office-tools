# PPMバブルチャート カスタムデータラベル実装プラン（XML直接操作）

## 背景・課題

openpyxlの`DataLabel`クラスには`tx`（カスタムテキスト）属性が存在しない。
`txPr`はテキスト書式のみで、テキスト内容は設定できない。
そのため、各バブルにセグメント名を表示するには、保存後のXMLを直接操作する必要がある。

## 対象ファイル

- `segment_analysis.py` — PPMバブルチャート生成処理

## 実装方針

openpyxlでExcelファイルを生成・保存した後、`lxml`でZIP内のチャートXMLを直接書き換える。

### OOXMLの`dLbl`構造（目標XML）

```xml
<c:dLbl>
  <c:idx val="0"/>
  <c:tx>
    <c:rich>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:t>セグメントA</a:t>
        </a:r>
      </a:p>
    </c:rich>
  </c:tx>
  <c:showLegendKey val="0"/>
  <c:showVal val="0"/>
  <c:showCatName val="0"/>
  <c:showSerName val="0"/>
  <c:showPercent val="0"/>
  <c:showBubbleSize val="0"/>
</c:dLbl>
```

**注意**: `<c:tx>` は `<c:idx>` の直後に配置する必要がある（OOXML仕様）。

---

## 実装手順

### Step 1: openpyxlでdLblを事前に配置する

XMLを後で書き換えるために、openpyxlで`DataLabel`を各データ点に対して作成しておく（`idx`だけ設定）。

```python
from openpyxl.chart.label import DataLabel, DataLabelList

def make_datalabel_list(n_points):
    """n_points個のDataLabelをidx=0,1,...で作成"""
    labels = []
    for i in range(n_points):
        lbl = DataLabel(idx=i)
        lbl.showLegendKey = False
        lbl.showVal = False
        lbl.showCatName = False
        lbl.showSerName = False
        lbl.showPercent = False
        lbl.showBubbleSize = False
        labels.append(lbl)

    dLbls = DataLabelList(dLbl=labels)
    dLbls.showLegendKey = False
    dLbls.showVal = False
    dLbls.showCatName = False
    dLbls.showSerName = False
    dLbls.showPercent = False
    dLbls.showBubbleSize = False
    return dLbls
```

### Step 2: ファイル保存後にXMLを書き換える関数

```python
import zipfile
import os
from lxml import etree

C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'c': C, 'a': A}

def inject_bubble_labels(xlsx_path, segment_names_per_series, chart_filename='xl/charts/chart1.xml'):
    """
    バブルチャートのdLblにカスタムテキスト（セグメント名）を注入する。

    Parameters
    ----------
    xlsx_path : str
        対象のExcelファイルパス
    segment_names_per_series : list[list[str]]
        シリーズごとのセグメント名リスト。
        例: [["セグメントA", "セグメントB"], ["セグメントC"]]
        シリーズが1つなら: [["A", "B", "C", ...]]
    chart_filename : str
        ZIP内のチャートXMLパス（複数チャートある場合は要調整）
    """
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        chart_xml = z.read(chart_filename)

    root = etree.fromstring(chart_xml)

    # 全シリーズを取得（bubbleChartのser要素）
    series_list = root.findall('.//c:bubbleChart/c:ser', NS)

    for ser_idx, ser in enumerate(series_list):
        if ser_idx >= len(segment_names_per_series):
            break
        seg_names = segment_names_per_series[ser_idx]

        dLbls = ser.find('c:dLbls', NS)
        if dLbls is None:
            continue

        for dlbl in dLbls.findall('c:dLbl', NS):
            idx_el = dlbl.find('c:idx', NS)
            if idx_el is None:
                continue
            idx = int(idx_el.get('val'))
            if idx >= len(seg_names):
                continue

            # <c:tx> 要素を構築
            tx = etree.Element(f'{{{C}}}tx')
            rich = etree.SubElement(tx, f'{{{C}}}rich')
            etree.SubElement(rich, f'{{{A}}}bodyPr')
            etree.SubElement(rich, f'{{{A}}}lstStyle')
            p = etree.SubElement(rich, f'{{{A}}}p')
            r = etree.SubElement(p, f'{{{A}}}r')
            t = etree.SubElement(r, f'{{{A}}}t')
            t.text = seg_names[idx]

            # idx の直後に挿入
            idx_el.addnext(tx)

    # ZIP書き戻し
    tmp_path = xlsx_path + '.tmp'
    with zipfile.ZipFile(xlsx_path, 'r') as zin, \
         zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == chart_filename:
                new_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item.filename))
    os.replace(tmp_path, xlsx_path)
```

### Step 3: 呼び出し側の変更

`segment_analysis.py` の `_create_ppm_analysis_sheet` 内で以下を変更する。

#### 3-1: `_append_data_section` の戻り値から `vcols` を受け取る

現状（L2375、L2385）の変数受け取りを修正し、`vcols` を保持する：

```python
# 変更前
lat_start, lat_end, lat_max_col, _, lat_chart_max_col = _append_data_section(LATEST_IDX, lat_sec_start)

# 変更後
lat_start, lat_end, lat_max_col, lat_vcols, lat_chart_max_col = _append_data_section(LATEST_IDX, lat_sec_start)
```

```python
# 変更前
five_start, five_end, _, _, five_chart_max_col = _append_data_section(FIVE_AGO_IDX, five_sec_start)

# 変更後
five_start, five_end, _, five_vcols, five_chart_max_col = _append_data_section(FIVE_AGO_IDX, five_sec_start)
```

#### 3-2: セグメント名をヘッダー行から取得する関数

```python
def _get_chart_segment_names(sec_start, vcols, chart_max_col):
    """
    データセクションのヘッダー行からチャート対象セグメント名リストを返す。
    vcols: _append_data_section が返す全列リスト
    chart_max_col: _add_series に渡したチャート用最大列（hokoku_col以左）
    """
    n_chart = chart_max_col - COL_DATA + 1  # チャートに含まれる列数
    names = []
    for k in range(n_chart):
        cell_val = ppm_ws.cell(sec_start, COL_DATA + k).value or ''
        names.append(str(cell_val))
    return names
```

#### 3-3: `wb.save(output_path)` の直後に注入処理を追加

チャートは `add_chart` の呼び出し順に `chart1.xml`, `chart2.xml`, ... と採番される。
`chart_latest` が先（L2423）、`chart_5y` が後（L2425）なので：

```python
wb.save(output_path)

# --- カスタムデータラベル注入 ---
lat_seg_names = _get_chart_segment_names(lat_start, lat_vcols, lat_chart_max_col)
inject_bubble_labels(output_path, [lat_seg_names], chart_filename='xl/charts/chart1.xml')

if chart_5y:
    five_seg_names = _get_chart_segment_names(five_start, five_vcols, five_chart_max_col)
    inject_bubble_labels(output_path, [five_seg_names], chart_filename='xl/charts/chart2.xml')
```

---

## 注意点・確認事項

1. **チャートXMLのパス**: このプログラムが生成するチャートはバブルチャート2つのみ。`chart_latest` が先（L2423）に `add_chart` されるため `chart1.xml`、`chart_5y` が後（L2425）のため `chart2.xml` と確定する。

2. **シリーズとデータ点の順序**: `DataLabel`の`idx`はシリーズ内のデータ点インデックス（0始まり）。
   `_add_series` でシリーズに渡す列順（`COL_DATA` から `chart_max_col` まで）と `segment_names` の順番を一致させること。

3. **シリーズは1本**: `_add_series` は1シリーズのみ追加するため、`segment_names_per_series` は常に要素1個のリスト `[names]` で渡す。

4. **既存のdLbl要素がない場合**: Step 1でopenpyxlから`DataLabelList`を設定しておかないと、
   XMLにdLbl要素が存在せずStep 2の注入が空振りになる。

5. **`lxml`の依存**: `requirements.txt`に`lxml`が含まれているか確認。

---

## テスト方法

1. 実装後にExcelを開き、バブルにセグメント名が表示されることを確認
2. ファイルを上書き保存してExcelが壊れないことを確認（OOXMLのXML宣言、エンコーディングの整合性）
3. `chart1.xml`をZIPから直接取り出してXML構造を目視確認する場合:
   ```bash
   python3 -c "
   import zipfile
   with zipfile.ZipFile('output.xlsx') as z:
       print(z.read('xl/charts/chart1.xml').decode())
   " | grep -A 20 'dLbl'
   ```
