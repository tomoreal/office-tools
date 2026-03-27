"""
Segment Analysis Module for XBRL to Excel Conversion

セグメント情報の分析シートを生成するモジュール
"""

import unicodedata


def add_segment_analysis_sheets(workbook, segment_sheets_info, debug_log=None):
    """
    セグメント情報シートから分析シートを生成

    Args:
        workbook: openpyxlワークブック
        segment_sheets_info: セグメントシート情報のリスト
            各要素は辞書で以下のキーを持つ:
            - sheet_name: シート名
            - ordered_keys: プレゼンテーションツリーの要素リスト [(full_path, pref_label), ...]
            - all_years_data: 全期間のデータ {role: {full_path: {(std, dim, period): value}}}
            - role: ロールURI
            - sorted_role_cols: ソート済みカラム情報のリスト
            - role_columns: ロールのカラム情報セット
            - current_standard: 現在の会計基準
            - segment_dict: セグメント用語辞書
            - common_dict: 共通用語辞書
            - labels_map: ラベルマップ
            - used_sheet_names: 使用済みシート名のセット
        debug_log: デバッグログ関数（オプション）
    """
    # デバッグログ関数がない場合はダミー関数を使用
    if debug_log is None:
        def debug_log(msg):
            pass

    debug_log(f"[Segment Analysis] Starting segment analysis sheet generation for {len(segment_sheets_info)} sheets")

    for info in segment_sheets_info:
        _create_segment_analysis_sheet(
            workbook=workbook,
            sheet_name=info['sheet_name'],
            ordered_keys=info['ordered_keys'],
            all_years_data=info['all_years_data'],
            role=info['role'],
            sorted_role_cols=info['sorted_role_cols'],
            role_columns=info['role_columns'],
            current_standard=info['current_standard'],
            segment_dict=info['segment_dict'],
            common_dict=info['common_dict'],
            labels_map=info['labels_map'],
            used_sheet_names=info['used_sheet_names'],
            debug_log=debug_log
        )

        # Create PPM analysis sheet for Japanese GAAP only
        if '日本基準' in info['sheet_name']:
            _create_ppm_analysis_sheet(
                workbook=workbook,
                analysis_sheet_name=info['sheet_name'] + '_分析',
                used_sheet_names=info['used_sheet_names'],
                debug_log=debug_log
            )


def _create_segment_analysis_sheet(workbook, sheet_name, ordered_keys, all_years_data,
                                   role, sorted_role_cols, role_columns, current_standard,
                                   segment_dict, common_dict, labels_map, used_sheet_names,
                                   debug_log):
    """
    セグメント分析シートを作成（内部関数）

    Args:
        workbook: openpyxlワークブック
        sheet_name: 元のシート名
        ordered_keys: プレゼンテーションツリーの要素リスト
        all_years_data: 全期間のデータ
        role: ロールURI
        sorted_role_cols: ソート済みカラム情報
        role_columns: ロールのカラム情報
        current_standard: 現在の会計基準
        segment_dict: セグメント用語辞書
        common_dict: 共通用語辞書
        labels_map: ラベルマップ
        used_sheet_names: 使用済みシート名のセット
        debug_log: デバッグログ関数
    """
    # 分析シート名を生成
    analysis_sheet_name = sheet_name + "_分析"
    if len(analysis_sheet_name) > 31:
        # Excel sheet name limit is 31 characters
        analysis_sheet_name = sheet_name[:28] + "_分析"

    debug_log(f"[Segment Analysis] Creating analysis sheet: {analysis_sheet_name}")

    # 分析シートを作成
    aws = workbook.create_sheet(title=analysis_sheet_name)
    used_sheet_names.add(analysis_sheet_name)

    # セグメントを横軸に配置（ユニークなディメンション）
    unique_dims = []
    for c in sorted_role_cols:
        # c is (std, dim, period)
        d = c[1] if len(c) == 3 else c[0]
        if d not in unique_dims:
            unique_dims.append(d)

    # All available years for this role (ascending)
    unique_periods = sorted(list(set(c[2] if len(c) == 3 else c[1] for c in role_columns)))

    debug_log(f"[Segment Analysis] Found {len(unique_dims)} segments and {len(unique_periods)} periods")

    # ヘッダー行を作成
    aws.append(["勘定科目", "年度"] + unique_dims)

    seen_rows_analysis = set()

    # データ行を作成
    for full_path_data in ordered_keys:
        full_path, pref_label = full_path_data
        el = full_path.split('::')[-1]
        if '|' in el:
            el = el.split('|')[0]

        # Skip irrelevant element types
        if el.endswith(("TextBlock", "Abstract", "Axis", "Member", "Table")):
            continue

        parts = el.split('_')
        base_name = parts[-1] if len(parts) > 1 else el

        # ラベルを取得
        if base_name in segment_dict:
            label = segment_dict[base_name]
        elif base_name in common_dict:
            label = common_dict[base_name]
        else:
            label = labels_map.get(el)
            if label:
                label = label.replace(' [メンバー]', '').replace(' [要素]', '').replace(' [区分]', '').strip()

        if not label:
            label = _convert_camel_case_to_title(base_name)

        # インデントを計算
        depth = len(full_path.split('::')) - 1
        indent_prefix = "　" * depth

        # ラベルから不要な接尾辞を削除
        display_label = label
        display_label = display_label.replace(' [目次項目]', '').replace(' [タイトル項目]', '')
        display_label = display_label.replace('（IFRS）', '').replace('(IFRS)', '')
        display_label = display_label.replace('、経営指標等', '')
        display_label = display_label.replace('、流動資産', '').replace('、非流動資産', '')
        display_label = display_label.replace('、流動負債', '').replace('、非流動負債', '')
        display_label = display_label.strip()

        # 各年度ごとに行を作成
        for period in unique_periods:
            row_data_analysis = [indent_prefix + display_label, period]
            has_numeric_data_analysis = False
            has_data_analysis = False

            for dim in unique_dims:
                # Search for (any_std, dim, period) - usually current_standard
                found_v = ""
                stds_to_check = [current_standard] if current_standard != 'JP_ALL' else ['IFRS', 'JP', 'US', 'JMIS']
                for s in stds_to_check:
                    v = all_years_data[role][full_path].get((s, dim, period))
                    if v is not None:
                        found_v = v
                        break
                val = found_v

                # 数値データかどうかをチェック
                if val:
                    val_clean = unicodedata.normalize('NFKC', str(val)).replace(',', '').strip()
                    try:
                        if val_clean and not any(c.isalpha() for c in val_clean):
                            val = float(val_clean)
                            has_numeric_data_analysis = True
                    except Exception:
                        pass

                row_data_analysis.append(val)
                if val != "":
                    has_data_analysis = True

            # データがある場合のみ行を追加
            if has_data_analysis:
                if not has_numeric_data_analysis:
                    continue

                # 重複チェック
                row_values_tuple = tuple(row_data_analysis[2:])
                row_key = (display_label, period, row_values_tuple)
                if row_key in seen_rows_analysis:
                    continue
                seen_rows_analysis.add(row_key)
                aws.append(row_data_analysis)

        # キャッシュ・フロー計算書の場合、特定の要素で停止
        if 'キャッシュ・フロー' in sheet_name and 'CashAndCashEquivalents' in el:
            if pref_label and pref_label.endswith(('periodEndLabel', 'totalLabel')):
                # Check if this is at natural end of hierarchy
                current_idx = ordered_keys.index(full_path_data)
                if current_idx >= len(ordered_keys) - 1:
                    break

                # Check if there are more CashAndCash items ahead
                has_more_cash_items = False
                for next_idx in range(current_idx + 1, len(ordered_keys)):
                    next_fp, _ = ordered_keys[next_idx]
                    next_el_name = next_fp.split('::')[-1]
                    if '|' in next_el_name:
                        next_el_name = next_el_name.split('|')[0]
                    if next_el_name.endswith(("Abstract", "TextBlock", "Table", "Axis", "Member")):
                        continue
                    # If we find another CashAndCash item, don't break yet
                    if 'CashAndCashEquivalents' in next_el_name:
                        has_more_cash_items = True
                        break
                    break

                if has_more_cash_items:
                    continue  # Don't break, process the next CashAndCash item

                # Check depth of next items if no more CashAndCash items
                current_depth = len(full_path.split('::')) - 1
                is_at_end = True
                for next_idx in range(current_idx + 1, len(ordered_keys)):
                    next_fp, _ = ordered_keys[next_idx]
                    next_el_name = next_fp.split('::')[-1]
                    if '|' in next_el_name:
                        next_el_name = next_el_name.split('|')[0]
                    if next_el_name.endswith(("Abstract", "TextBlock", "Table", "Axis", "Member")):
                        continue
                    next_depth = len(next_fp.split('::')) - 1
                    if next_depth > current_depth:
                        is_at_end = False
                    break
                if is_at_end:
                    break

    # セル書式を適用
    for row in aws.iter_rows(min_row=2, max_row=aws.max_row, min_col=3, max_col=aws.max_column):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = r'#,##0_ ;[Red]\-#,##0 '

    # ウィンドウ枠の固定 (B2で固定: A列と1行目を固定)
    aws.freeze_panes = 'B2'

    # 列幅の設定
    # A列: 31
    aws.column_dimensions['A'].width = 31
    # B列以降: 12
    for col_idx in range(2, aws.max_column + 1):
        col_letter = chr(64 + col_idx)  # B=66, C=67, etc.
        aws.column_dimensions[col_letter].width = 12

    debug_log(f"[Segment Analysis] Completed analysis sheet: {analysis_sheet_name} with {aws.max_row - 1} data rows")


def _convert_camel_case_to_title(name):
    """
    キャメルケースをタイトルケースに変換

    Args:
        name: 変換する文字列

    Returns:
        変換後の文字列
    """
    import re
    # Insert space before uppercase letters
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1 \2', name)
    # Insert space before uppercase letters that follow lowercase letters or numbers
    return re.sub('([a-z0-9])([A-Z])', r'\1 \2', s1)


def _create_ppm_analysis_sheet(workbook, analysis_sheet_name, used_sheet_names, debug_log):
    """
    PPM分析シートを作成（内部関数）

    Args:
        workbook: openpyxlワークブック
        analysis_sheet_name: 参照元の分析シート名
        used_sheet_names: 使用済みシート名のセット
        debug_log: デバッグログ関数
    """
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BubbleChart, Reference, Series
    from openpyxl.chart.label import DataLabelList

    # PPM分析シート名を生成
    ppm_sheet_name = analysis_sheet_name + "_PPM分析用"
    if len(ppm_sheet_name) > 31:
        # Excel sheet name limit is 31 characters
        ppm_sheet_name = analysis_sheet_name[:18] + "_PPM分析用"

    debug_log(f"[PPM Analysis] Creating PPM analysis sheet: {ppm_sheet_name}")

    # Check if analysis sheet exists
    if analysis_sheet_name not in workbook.sheetnames:
        debug_log(f"[PPM Analysis] Analysis sheet '{analysis_sheet_name}' not found, skipping PPM sheet")
        return

    # Get the analysis sheet
    analysis_ws = workbook[analysis_sheet_name]

    # PPM分析シートを作成
    ppm_ws = workbook.create_sheet(title=ppm_sheet_name)
    used_sheet_names.add(ppm_sheet_name)

    # Escape single quotes in sheet name for formula
    escaped_sheet_name = analysis_sheet_name.replace("'", "''")

    # Get the number of columns in the analysis sheet
    max_col = analysis_ws.max_column

    # Row 1: Header row (references from analysis sheet row 1)
    header_row = []
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        formula = f"=IF('{escaped_sheet_name}'!{col_letter}1=\"\",\"\",'{escaped_sheet_name}'!{col_letter}1)"
        header_row.append(formula)
    ppm_ws.append(header_row)

    # Track row numbers for numeric data (Sales and Segment Profit)
    sales_start_row = ppm_ws.max_row + 1  # Track where sales rows start

    # Rows 2-12: 売上 (Sales) - references rows 24-34 from analysis sheet
    for src_row in range(24, 35):  # 24-34 inclusive (11 rows)
        data_row = ["　売上"]  # Fixed account name in column A
        for col_idx in range(2, max_col + 1):
            col_letter = get_column_letter(col_idx)
            formula = f"=IF('{escaped_sheet_name}'!{col_letter}{src_row}=\"\",\"\",'{escaped_sheet_name}'!{col_letter}{src_row})"
            data_row.append(formula)
        ppm_ws.append(data_row)

    sales_end_row = ppm_ws.max_row  # Track where sales rows end
    profit_start_row = ppm_ws.max_row + 1  # Track where profit rows start

    # Rows 13-23: セグメント利益 (Segment Profit) - references rows 35-45 from analysis sheet
    for src_row in range(35, 46):  # 35-45 inclusive (11 rows)
        data_row = ["　セグメント利益"]  # Fixed account name in column A
        for col_idx in range(2, max_col + 1):
            col_letter = get_column_letter(col_idx)
            formula = f"=IF('{escaped_sheet_name}'!{col_letter}{src_row}=\"\",\"\",'{escaped_sheet_name}'!{col_letter}{src_row})"
            data_row.append(formula)
        ppm_ws.append(data_row)

    profit_end_row = ppm_ws.max_row  # Track where profit rows end

    # Empty row separator
    ppm_ws.append([""] * max_col)

    # Track growth rate rows
    growth_start_row = ppm_ws.max_row + 1  # Track where growth rate rows start

    # 売上高対前年増加率 (Sales YoY Growth Rate)
    # Calculate number of data rows (should be 11 rows based on sales data)
    num_data_rows = sales_end_row - sales_start_row + 1
    for idx in range(num_data_rows):  # For each year
        growth_row = ["売上高対前年増加率"]
        # Calculate which row in sales data to reference
        year_row_ref = sales_start_row + idx  # Reference to corresponding sales row

        for col_idx in range(2, max_col + 1):
            col_letter = get_column_letter(col_idx)
            if col_idx == 2:  # Column B: reference the year from sales data
                formula = f"=B{year_row_ref}"
            elif idx == 0:  # First year: empty data columns (no previous year to compare)
                formula = ""
            else:  # Calculate growth rate
                # Current year is in row (sales_start_row + idx), previous year is in row (sales_start_row + idx - 1)
                current_row_ref = sales_start_row + idx
                previous_row_ref = sales_start_row + idx - 1
                formula = f"=IF(OR({col_letter}{current_row_ref}=\"\",{col_letter}{previous_row_ref}=\"\"),\"\",{col_letter}{current_row_ref}/{col_letter}{previous_row_ref}-1)"
            growth_row.append(formula)
        ppm_ws.append(growth_row)

    growth_end_row = ppm_ws.max_row  # Track where growth rate rows end

    # Empty row separator
    ppm_ws.append([""] * max_col)

    # Track profit margin rows
    margin_start_row = ppm_ws.max_row + 1  # Track where profit margin rows start

    # 売上高利益率 (Sales Profit Margin)
    for idx in range(num_data_rows):  # For each year
        margin_row = ["売上高利益率"]
        # Calculate which row to reference
        year_row_ref = sales_start_row + idx  # Reference to corresponding sales row

        for col_idx in range(2, max_col + 1):
            col_letter = get_column_letter(col_idx)
            if col_idx == 2:  # Column B: reference the year from sales data
                formula = f"=B{year_row_ref}"
            else:  # Columns C onwards: Calculate profit margin
                # Sales is in sales rows, Segment profit is in profit rows
                sales_row_ref = sales_start_row + idx
                profit_row_ref = profit_start_row + idx
                formula = f"=IF(OR({col_letter}{profit_row_ref}=\"\",{col_letter}{sales_row_ref}=\"\"),\"\",{col_letter}{profit_row_ref}/{col_letter}{sales_row_ref})"
            margin_row.append(formula)
        ppm_ws.append(margin_row)

    margin_end_row = ppm_ws.max_row  # Track where profit margin rows end

    # Apply formatting
    # Freeze panes at B2
    ppm_ws.freeze_panes = 'B2'

    # Set column widths
    ppm_ws.column_dimensions['A'].width = 31
    for col_idx in range(2, max_col + 1):
        col_letter = get_column_letter(col_idx)
        ppm_ws.column_dimensions[col_letter].width = 12

    # Set number format for data rows (Sales and Segment Profit)
    # Format: #,##0_ ;[Red]\-#,##0 (thousand separators, negative in red)
    for row_idx in range(sales_start_row, profit_end_row + 1):  # All Sales and Profit rows
        for col_idx in range(3, max_col + 1):  # Starting from column C (data columns)
            col_letter = get_column_letter(col_idx)
            cell = ppm_ws[f'{col_letter}{row_idx}']
            cell.number_format = r'#,##0_ ;[Red]\-#,##0 '

    # Set number format for percentage rows (growth rate and profit margin)
    from openpyxl.styles import numbers
    for row_idx in list(range(growth_start_row, growth_end_row + 1)) + list(range(margin_start_row, margin_end_row + 1)):
        for col_idx in range(3, max_col + 1):  # Starting from column C (data columns)
            col_letter = get_column_letter(col_idx)
            cell = ppm_ws[f'{col_letter}{row_idx}']
            cell.number_format = '0%'  # Display as percentage (0% format like in sample)

    # ===== Create consolidated data area for bubble chart =====
    # Empty row separators (two rows for spacing)
    ppm_ws.append([""] * max_col)
    ppm_ws.append([""] * max_col)

    # Data consolidation starts here (transposed format for chart)
    data_start_row = ppm_ws.max_row + 1

    # Find the column containing "報告セグメント" in the header row (row 1)
    # This will be the last column for chart data
    # We want the column that has JUST "報告セグメント" or ends with it (not combined columns)
    chart_end_col = max_col
    for col_idx in range(3, max_col + 1):
        header_val = analysis_ws.cell(1, col_idx).value
        if header_val:
            header_str = str(header_val).strip()
            # Look for the column that contains "報告セグメント" without "以外" or "その他"
            # This should match columns like "報告セグメント", "報告セグメント合計" etc.
            if "報告セグメント" in header_str and "以外" not in header_str:
                chart_end_col = col_idx
                break

    # Row 1: Year (A column reference) and Segment names
    # A50: Year reference from latest sales data
    # B50: "セグメント名"
    # C50-J50: Segment names from header row
    consolidated_header = [f"=B{sales_end_row}", "セグメント名"]
    for col_idx in range(3, chart_end_col + 1):
        col_letter = get_column_letter(col_idx)
        consolidated_header.append(f"={col_letter}1")
    # Replace last column header with "計"
    consolidated_header[-1] = "計"
    ppm_ws.append(consolidated_header)

    # Get the year value from the source for chart title
    # The year is in analysis_ws column B, row sales_end_row
    year_value_for_title = analysis_ws.cell(sales_end_row, 2).value

    # Row 2: Profit margin (売上高利益率)
    # A: Year reference, B: Label, C-J: Profit margin values
    margin_data_row = [f"=B{margin_end_row}", f"=A{margin_end_row}"]
    for col_idx in range(3, chart_end_col + 1):
        col_letter = get_column_letter(col_idx)
        margin_data_row.append(f"={col_letter}{margin_end_row}")
    ppm_ws.append(margin_data_row)

    # Row 3: Growth rate (売上高対前年増加率)
    # A: Year reference, B: Label, C-J: Growth rate values
    growth_data_row = [f"=B{growth_end_row}", f"=A{growth_end_row}"]
    for col_idx in range(3, chart_end_col + 1):
        col_letter = get_column_letter(col_idx)
        growth_data_row.append(f"={col_letter}{growth_end_row}")
    ppm_ws.append(growth_data_row)

    # Row 4: Sales (売上)
    # A: Year reference, B: Label (TRIM), C-J: Sales values
    sales_data_row = [f"=B{sales_end_row}", f"=TRIM(A{sales_end_row})"]
    for col_idx in range(3, chart_end_col):  # Up to chart_end_col - 1 (not including last column)
        col_letter = get_column_letter(col_idx)
        sales_data_row.append(f"={col_letter}{sales_end_row}")
    # Last column: Scale down by 1% for better chart display
    last_col_letter = get_column_letter(chart_end_col)
    sales_data_row.append(f"={last_col_letter}{sales_end_row}*1%")
    ppm_ws.append(sales_data_row)

    data_end_row = ppm_ws.max_row

    # Apply formatting to consolidated data area
    # Row with profit margin: percentage format
    for col_idx in range(3, chart_end_col + 1):
        col_letter = get_column_letter(col_idx)
        cell = ppm_ws[f'{col_letter}{data_start_row + 1}']
        cell.number_format = '0%'

    # Row with growth rate: percentage format
    for col_idx in range(3, chart_end_col + 1):
        col_letter = get_column_letter(col_idx)
        cell = ppm_ws[f'{col_letter}{data_start_row + 2}']
        cell.number_format = '0%'

    # Row with sales: thousand separator format
    for col_idx in range(3, chart_end_col + 1):
        col_letter = get_column_letter(col_idx)
        cell = ppm_ws[f'{col_letter}{data_start_row + 3}']
        cell.number_format = r'#,##0_);[Red](#,##0)'

    # ===== Create Bubble Chart =====
    chart = BubbleChart()
    chart.style = 2  # Use a predefined style

    # Set chart title dynamically based on the latest year
    # Use the year_value_for_title we got earlier when creating the consolidated header
    if year_value_for_title:
        # Format as YYYY/MM
        if isinstance(year_value_for_title, str):
            # It's already a string, try to parse it
            import datetime
            try:
                if '-' in year_value_for_title:
                    date_obj = datetime.datetime.strptime(year_value_for_title, '%Y-%m-%d')
                    year_str = date_obj.strftime('%Y/%m')
                else:
                    year_str = year_value_for_title[:7].replace('-', '/')  # Simple conversion
            except:
                year_str = ""
        elif hasattr(year_value_for_title, 'strftime'):
            # It's a datetime object
            year_str = year_value_for_title.strftime('%Y/%m')
        else:
            year_str = ""
    else:
        year_str = ""

    chart.title = f"PPM分析 {year_str}"

    # Configure chart
    chart.height = 15
    chart.width = 15

    # Configure X-axis (売上高利益率)
    chart.x_axis.title = "売上高利益率"
    #chart.x_axis.majorGridlines = None  # Remove gridlines if needed

    # Configure Y-axis (売上高対前年増加率)
    chart.y_axis.title = "売上高対前年増加率"
    #chart.y_axis.majorGridlines = None  # Remove gridlines if needed

    # 目盛ラベル表示
    chart.x_axis.tickLblPos = "nextTo"
    chart.y_axis.tickLblPos = "nextTo"

    # 軸自体も明示的に有効化
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    # Create data series
    # X values: Profit margin (row data_start_row + 1, columns C to chart_end_col)
    # Y values: Growth rate (row data_start_row + 2, columns C to chart_end_col)
    # Bubble size: Sales (row data_start_row + 3, columns C to chart_end_col)

    xvalues = Reference(ppm_ws, min_col=3, min_row=data_start_row + 1, max_col=chart_end_col, max_row=data_start_row + 1)
    yvalues = Reference(ppm_ws, min_col=3, min_row=data_start_row + 2, max_col=chart_end_col, max_row=data_start_row + 2)
    size = Reference(ppm_ws, min_col=3, min_row=data_start_row + 3, max_col=chart_end_col, max_row=data_start_row + 3)

    # For bubble charts, the Series API is: Series(values=yvalues, xvalues=xvalues, zvalues=size)
    series = Series(values=yvalues, xvalues=xvalues, zvalues=size, title="")

    # Add data labels showing segment names
    # Create individual data labels for each bubble with segment names
    from openpyxl.chart.label import DataLabel
    from openpyxl.chart.text import Text
    from openpyxl.chart.data_source import StrRef

    chart.legend = None # hide legend

    chart.series.append(series)

    # Position chart starting at column B, below the data
    # chart_row = data_end_row + 1
    # ppm_ws.add_chart(chart, f'B{chart_row}')

    # ===== Create 5-year-old data section =====
    # Add blank row separator
    ppm_ws.append([""] * max_col)

    # Find the row for 5 years ago (6th row from the end in sales data, since last row is latest)
    # Sales data is in rows sales_start_row to sales_end_row (11 rows for 11 years)
    # 5 years ago would be at: sales_end_row - 5
    five_years_ago_offset = 5
    if (sales_end_row - sales_start_row + 1) > five_years_ago_offset:
        five_year_sales_row = sales_end_row - five_years_ago_offset
        five_year_profit_row = profit_end_row - five_years_ago_offset
        five_year_growth_row = growth_end_row - five_years_ago_offset
        five_year_margin_row = margin_end_row - five_years_ago_offset

        # Data consolidation for 5-year-old data starts here
        five_year_data_start_row = ppm_ws.max_row + 1

        # Get the year value from 5 years ago for chart title
        year_value_five_years_ago = analysis_ws.cell(five_year_sales_row, 2).value

        # Row 1: Year and Segment names
        five_year_header = [f"=B{five_year_sales_row}", "セグメント名"]
        for col_idx in range(3, chart_end_col + 1):
            col_letter = get_column_letter(col_idx)
            five_year_header.append(f"={col_letter}1")
        # Replace last column header with "計"
        five_year_header[-1] = "計"
        ppm_ws.append(five_year_header)

        # Row 2: Profit margin (売上高利益率)
        five_year_margin_data = [f"=B{five_year_margin_row}", f"=A{five_year_margin_row}"]
        for col_idx in range(3, chart_end_col + 1):
            col_letter = get_column_letter(col_idx)
            five_year_margin_data.append(f"={col_letter}{five_year_margin_row}")
        ppm_ws.append(five_year_margin_data)

        # Row 3: Growth rate (売上高対前年増加率)
        five_year_growth_data = [f"=B{five_year_growth_row}", f"=A{five_year_growth_row}"]
        for col_idx in range(3, chart_end_col + 1):
            col_letter = get_column_letter(col_idx)
            five_year_growth_data.append(f"={col_letter}{five_year_growth_row}")
        ppm_ws.append(five_year_growth_data)

        # Row 4: Sales (売上)
        five_year_sales_data = [f"=B{five_year_sales_row}", f"=TRIM(A{five_year_sales_row})"]
        for col_idx in range(3, chart_end_col):
            col_letter = get_column_letter(col_idx)
            five_year_sales_data.append(f"={col_letter}{five_year_sales_row}")
        # Last column: Scale down by 1%
        five_year_sales_data.append(f"={last_col_letter}{five_year_sales_row}*1%")
        ppm_ws.append(five_year_sales_data)

        five_year_data_end_row = ppm_ws.max_row

        # Apply formatting to 5-year-old data area
        # Row with profit margin: percentage format
        for col_idx in range(3, chart_end_col + 1):
            col_letter = get_column_letter(col_idx)
            cell = ppm_ws[f'{col_letter}{five_year_data_start_row + 1}']
            cell.number_format = '0%'

        # Row with growth rate: percentage format
        for col_idx in range(3, chart_end_col + 1):
            col_letter = get_column_letter(col_idx)
            cell = ppm_ws[f'{col_letter}{five_year_data_start_row + 2}']
            cell.number_format = '0%'

        # Row with sales: thousand separator format
        for col_idx in range(3, chart_end_col + 1):
            col_letter = get_column_letter(col_idx)
            cell = ppm_ws[f'{col_letter}{five_year_data_start_row + 3}']
            cell.number_format = r'#,##0_);[Red](#,##0)'

        # ===== Create Bubble Chart for 5-year-old data =====
        chart_5y = BubbleChart()
        chart_5y.style = 2

        # Set chart title for 5 years ago
        if year_value_five_years_ago:
            if isinstance(year_value_five_years_ago, str):
                import datetime
                try:
                    if '-' in year_value_five_years_ago:
                        date_obj = datetime.datetime.strptime(year_value_five_years_ago, '%Y-%m-%d')
                        year_str_5y = date_obj.strftime('%Y/%m')
                    else:
                        year_str_5y = year_value_five_years_ago[:7].replace('-', '/')
                except:
                    year_str_5y = ""
            elif hasattr(year_value_five_years_ago, 'strftime'):
                year_str_5y = year_value_five_years_ago.strftime('%Y/%m')
            else:
                year_str_5y = ""
        else:
            year_str_5y = ""

        chart_5y.title = f"PPM分析 {year_str_5y}"

        # Configure chart
        chart_5y.height = 15
        chart_5y.width = 15

        # Configure axes
        chart_5y.x_axis.title = "売上高利益率"
        chart_5y.y_axis.title = "売上高対前年増加率"
        chart_5y.x_axis.tickLblPos = "nextTo"
        chart_5y.y_axis.tickLblPos = "nextTo"
        chart_5y.x_axis.delete = False
        chart_5y.y_axis.delete = False

        # Create data series for 5-year-old data
        xvalues_5y = Reference(ppm_ws, min_col=3, min_row=five_year_data_start_row + 1, max_col=chart_end_col, max_row=five_year_data_start_row + 1)
        yvalues_5y = Reference(ppm_ws, min_col=3, min_row=five_year_data_start_row + 2, max_col=chart_end_col, max_row=five_year_data_start_row + 2)
        size_5y = Reference(ppm_ws, min_col=3, min_row=five_year_data_start_row + 3, max_col=chart_end_col, max_row=five_year_data_start_row + 3)

        series_5y = Series(values=yvalues_5y, xvalues=xvalues_5y, zvalues=size_5y, title="")
        chart_5y.legend = None
        chart_5y.series.append(series_5y)

        # Position chart starting at column B, below the data
        chart_row = five_year_data_end_row + 2
        ppm_ws.add_chart(chart, f'B{chart_row}')

        # Position the 5-year chart to the right of the current chart
        # Current chart is at column B (column 2), width 15 columns
        # Position new chart at column Q (column 17) to give some spacing
        chart_5y_col = 'I'
        ppm_ws.add_chart(chart_5y, f'{chart_5y_col}{chart_row}')

        debug_log(f"[PPM Analysis] Added 5-year-old data section and chart")

    debug_log(f"[PPM Analysis] Completed PPM analysis sheet: {ppm_sheet_name} with {ppm_ws.max_row} rows and bubble chart(s)")
