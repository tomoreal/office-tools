現在、セグメント情報は、最新から過去にさかのぼって、１年分ずつを保持しています。
一番新しい年度以外は、前年度の値を使っています。
この方式だと、セグメント変更があった場合、正しいPPMグラフを作れません。
各年度、当年度と前年度のデータを保持していれば、必ずセグメント数が一致するので正しいPPMグラフが作れます。
これは、可能でしょうか？

検討結果
現在の仕組みと問題点

年度:  2020   2021   2022   2023   2024   2025
       ↑      ↑      ↑      ↑      ↑      ↑
      各ZIPから1年分ずつ取得（それぞれ別ファイルのデータ）
PPM成長率の計算: value(N) / value(N-1) - 1
→ N期のデータ と N-1期のデータ が別々の有報から来る

セグメント組み替えが起きた場合、N-1期のセグメント構成がN期と異なるため、成長率が正確でない。

提案方式の効果
各有報は 当期 と 前期 の2期分を掲載している。

同じ有報の前期データを使えば、組み替え後の組み替え前比較データが使える。

2025年有報:  [当期2025, 前期2024(組み替え後)]  ← 同一セグメント定義
2024年有報:  [当期2024(旧定義), 前期2023(旧定義)]
PPM成長率に 同一有報の当期/前期 を使えば常にセグメント一致が保証される。

実現可能か
技術的には可能。ただし変更規模は大きい。

課題1: XBRL抽出段階の変更が必要
現在、XBRLから年度別データを取り出して分析シートに積み上げる際、どの有報から来た前期データか が保持されていない。

変更が必要な箇所:

convert_xbrl_to_excel.py の XBRL パース部分（CurrentYearDuration / Prior1YearDuration コンテキストの対応関係を保持）
segment_analysis.py の _create_segment_analysis_sheet() でデータ構造変更
課題2: 分析シートの構造変更が必要
提案方式では、各年度のデータに「どの有報の当期/前期か」の対応が必要になる。


案B: PPM専用の前期対応辞書を保持（分析シートは変更しない）

# XBRLパース時に別途構築
prev_lookup: dict[(label, year_N)] → value_from_year_N_report_prior_period
→ 分析シートへの影響なし。PPM関数にこの辞書を渡す形で実装可能


推奨アプローチ（もし実装するなら）
案B（分析シート非破壊）が現実的：

XBRL処理時に「各年度の当期・前期ペア」を辞書として別途保持
PPM関数にこの辞書を追加引数として渡す
成長率計算で prev_lookup が存在する年度はそちらを優先使用
存在しない年度は現状通りのフォールバック動作

変更規模の目安：

convert_xbrl_to_excel.py の XBRL コンテキスト処理部分（大）
segment_analysis.py の PPM 関数（中）
他シートへの影響：なし（案B採用の場合）

この案Bで実装計画を立てて、mdファイルに保存してください。3/31の午前中に実装を行います。

Agent:Explore XBRL to analysis sheet data flow

In /home/tomo/work_office_edinet_api, I need to understand the data flow from XBRL parsing to the segment analysis sheet creation, specifically:

1. In convert_xbrl_to_excel.py: how is segment data extracted from XBRL? Look for where "all_years_data" or equivalent is built for segments. Show the relevant code sections (~30 lines context) around where segment data is collected per year. Include line numbers.

2. How is `_create_segment_analysis_sheet` (in segment_analysis.py) called from convert_xbrl_to_excel.py? What arguments are passed? Show the call site with context.

3. In segment_analysis.py's `_create_segment_analysis_sheet`, how is `all_years_data` structured? Show the first 50-80 lines of this function to understand the input data format.

4. In convert_xbrl_to_excel.py: how are `_create_ppm_analysis_sheet` and `_create_ppm_analysis_sheet_ifrs` called? What arguments are passed?

5. What is `sorted_role_cols` / `all_periods_data` structure in the segment analysis context?

Please be thorough and show actual code with line numbers.
Globpattern: "**/*.py"
Found 101 files
Bash
IN
find /home/tomo/work_office_edinet_api -name "*.py" -path "*/venv" -prune -o -name "convert_xbrl_to_excel.py" -o -name "segment_analysis.py" | grep -v venv

OUT
/home/tomo/work_office_edinet_api/segment_analysis.py
/home/tomo/work_office_edinet_api/convert_xbrl_to_excel.py

