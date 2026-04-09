"""
PPMバブルチャート カスタムデータラベル注入モジュール

openpyxlの DataLabel では <c:tx> によるカスタムテキストを設定できないため、
wb.save() 後に ZIP 内のチャート XML を lxml で直接書き換える。
"""

import os
import zipfile

from lxml import etree
from openpyxl.chart.label import DataLabel, DataLabelList

# OOXML 名前空間
_C   = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_C15 = 'http://schemas.microsoft.com/office/drawing/2012/chart'
_NS  = {'c': _C, 'a': _A, 'c15': _C15}


def make_datalabel_list(n_points):
    """
    n_points 個の DataLabel を idx=0,1,... で作成して DataLabelList を返す。

    openpyxl でシリーズに設定しておくことで、保存後の XML に <c:dLbls>/<c:dLbl>
    要素が生成される。これを inject_bubble_labels() で書き換える。
    """
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


def inject_bubble_labels(xlsx_path, segment_names_per_series, chart_filename):
    """
    保存済み xlsx のチャート XML に <c:tx><c:rich>...<a:t>セグメント名</a:t>
    を注入する（直接テキスト埋め込み方式）。

    Parameters
    ----------
    xlsx_path : str
        対象の Excel ファイルパス（保存済み）
    segment_names_per_series : list[list[str]]
        シリーズごとのセグメント名リスト。
        このチャートはシリーズ 1 本なので [[名前1, 名前2, ...]] の形式で渡す。
    chart_filename : str
        ZIP 内のチャート XML パス（例: 'xl/charts/chart1.xml'）
    """
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        chart_xml = z.read(chart_filename)

    root = etree.fromstring(chart_xml)

    # bubbleChart 内の全シリーズを取得
    series_list = root.findall('.//c:bubbleChart/c:ser', _NS)

    for ser_idx, ser in enumerate(series_list):
        if ser_idx >= len(segment_names_per_series):
            break
        seg_names = segment_names_per_series[ser_idx]

        dLbls = ser.find('c:dLbls', _NS)
        if dLbls is None:
            continue

        for dlbl in dLbls.findall('c:dLbl', _NS):
            idx_el = dlbl.find('c:idx', _NS)
            if idx_el is None:
                continue
            idx = int(idx_el.get('val'))
            if idx >= len(seg_names):
                continue

            # <c:tx><c:rich>...<a:t>セグメント名</a:t>...</c:rich></c:tx> を構築
            tx = etree.Element(f'{{{_C}}}tx')
            rich = etree.SubElement(tx, f'{{{_C}}}rich')
            etree.SubElement(rich, f'{{{_A}}}bodyPr')
            etree.SubElement(rich, f'{{{_A}}}lstStyle')
            p = etree.SubElement(rich, f'{{{_A}}}p')
            r = etree.SubElement(p, f'{{{_A}}}r')
            t = etree.SubElement(r, f'{{{_A}}}t')
            t.text = seg_names[idx]

            # OOXML 仕様: <c:tx> は <c:idx> の直後に配置する
            idx_el.addnext(tx)

        # dLbls レベルにリーダー線を有効化する要素を追加
        # <c:showLeaderLines val="1"/> — 旧形式（openpyxl が書いた既存要素があれば上書き）
        show_ll = dLbls.find('c:showLeaderLines', _NS)
        if show_ll is None:
            show_ll = etree.SubElement(dLbls, f'{{{_C}}}showLeaderLines')
        show_ll.set('val', '1')

        # dLbls の extLst に <c15:showLeaderLines val="1"/> を追加（Excel 2013+ 形式）
        extlst = dLbls.find('c:extLst', _NS)
        if extlst is None:
            extlst = etree.SubElement(dLbls, f'{{{_C}}}extLst')
        ext = extlst.find(f'c:ext[@uri="{{CE6537A1-D6FC-4f65-9D91-7224C49458BB}}"]', _NS)
        if ext is None:
            ext = etree.SubElement(extlst, f'{{{_C}}}ext')
            ext.set('uri', '{CE6537A1-D6FC-4f65-9D91-7224C49458BB}')
        c15_ll = ext.find('c15:showLeaderLines', _NS)
        if c15_ll is None:
            c15_ll = etree.SubElement(ext, f'{{{_C15}}}showLeaderLines')
        c15_ll.set('val', '1')

    # ZIP に書き戻し（tmp 経由でアトミックに置換）
    tmp_path = xlsx_path + '.tmp'
    with zipfile.ZipFile(xlsx_path, 'r') as zin, \
         zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == chart_filename:
                new_xml = etree.tostring(
                    root, xml_declaration=True, encoding='UTF-8', standalone=True
                )
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item.filename))

    os.replace(tmp_path, xlsx_path)
