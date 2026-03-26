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

    # Rows 2-12: 売上 (Sales) - references rows 24-34 from analysis sheet
    for src_row in range(24, 35):  # 24-34 inclusive (11 rows)
        data_row = ["　売上"]  # Fixed account name in column A
        for col_idx in range(2, max_col + 1):
            col_letter = get_column_letter(col_idx)
            formula = f"=IF('{escaped_sheet_name}'!{col_letter}{src_row}=\"\",\"\",'{escaped_sheet_name}'!{col_letter}{src_row})"
            data_row.append(formula)
        ppm_ws.append(data_row)

    # Rows 13-23: セグメント利益 (Segment Profit) - references rows 35-45 from analysis sheet
    for src_row in range(35, 46):  # 35-45 inclusive (11 rows)
        data_row = ["　セグメント利益"]  # Fixed account name in column A
        for col_idx in range(2, max_col + 1):
            col_letter = get_column_letter(col_idx)
            formula = f"=IF('{escaped_sheet_name}'!{col_letter}{src_row}=\"\",\"\",'{escaped_sheet_name}'!{col_letter}{src_row})"
            data_row.append(formula)
        ppm_ws.append(data_row)

    # Row 24: Empty row
    ppm_ws.append([""] * max_col)

    # Row 25-35: 売上高対前年増加率 (Sales YoY Growth Rate)
    # Rows 25-35: Each row references the corresponding year from B2-B12
    for ppm_row in range(25, 36):  # Rows 25-35 (11 rows)
        growth_row = ["売上高対前年増加率"]
        # Calculate which row in B2-B12 to reference
        # Row 25 -> B2, Row 26 -> B3, ..., Row 35 -> B12
        year_row_ref = ppm_row - 23  # Row 25 -> 2, Row 26 -> 3, etc.

        for col_idx in range(2, max_col + 1):
            col_letter = get_column_letter(col_idx)
            if col_idx == 2:  # Column B: reference the year from B2-B12
                formula = f"=B{year_row_ref}"
            elif ppm_row == 25:  # Row 25: empty data columns (first year has no growth rate)
                formula = ""
            else:  # Rows 26-35: Calculate growth rate
                # Current year is in row (ppm_row - 23), previous year is in row (ppm_row - 24)
                current_row_ref = ppm_row - 23  # Row 26 -> 3, Row 27 -> 4, etc.
                previous_row_ref = ppm_row - 24  # Row 26 -> 2, Row 27 -> 3, etc.
                formula = f"=IF(OR({col_letter}{current_row_ref}=\"\",{col_letter}{previous_row_ref}=\"\"),\"\",{col_letter}{current_row_ref}/{col_letter}{previous_row_ref}-1)"
            growth_row.append(formula)
        ppm_ws.append(growth_row)

    # Row 36: Empty row
    ppm_ws.append([""] * max_col)

    # Rows 37-47: 売上高利益率 (Sales Profit Margin)
    for ppm_row in range(37, 48):  # Rows 37-47 (11 rows)
        margin_row = ["売上高利益率"]
        # Calculate which row in B2-B12 to reference
        # Row 37 -> B2, Row 38 -> B3, ..., Row 47 -> B12
        year_row_ref = ppm_row - 35  # Row 37 -> 2, Row 38 -> 3, etc.

        for col_idx in range(2, max_col + 1):
            col_letter = get_column_letter(col_idx)
            if col_idx == 2:  # Column B: reference the year from B2-B12
                formula = f"=B{year_row_ref}"
            else:  # Columns C onwards: Calculate profit margin
                # Sales is in row (ppm_row - 35 + 2), Segment profit is in row (ppm_row - 35 + 13)
                sales_row_ref = ppm_row - 35  # Row 37 -> 2, Row 38 -> 3, etc.
                profit_row_ref = ppm_row - 24  # Row 37 -> 13, Row 38 -> 14, etc.
                formula = f"=IF(OR({col_letter}{profit_row_ref}=\"\",{col_letter}{sales_row_ref}=\"\"),\"\",{col_letter}{profit_row_ref}/{col_letter}{sales_row_ref})"
            margin_row.append(formula)
        ppm_ws.append(margin_row)

    # Apply formatting
    # Freeze panes at B2
    ppm_ws.freeze_panes = 'B2'

    # Set column widths
    ppm_ws.column_dimensions['A'].width = 31
    for col_idx in range(2, max_col + 1):
        col_letter = get_column_letter(col_idx)
        ppm_ws.column_dimensions[col_letter].width = 12

    # Set number format for percentage rows (rows 25-47 except row 36)
    from openpyxl.styles import numbers
    for row_idx in list(range(25, 36)) + list(range(37, 48)):  # Rows 25-35 and 37-47
        for col_idx in range(3, max_col + 1):  # Starting from column C (data columns)
            col_letter = get_column_letter(col_idx)
            cell = ppm_ws[f'{col_letter}{row_idx}']
            cell.number_format = '0%'  # Display as percentage (0% format like in sample)

    debug_log(f"[PPM Analysis] Completed PPM analysis sheet: {ppm_sheet_name} with {ppm_ws.max_row} rows")
