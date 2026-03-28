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
            # Use the same sheet name transformation as _create_segment_analysis_sheet
            # (replacing "注記" with "連結")
            base_name = info['sheet_name'].replace("注記", "連結")
            analysis_sheet_name = base_name + "_分析"
            if len(analysis_sheet_name) > 31:
                analysis_sheet_name = base_name[:28] + "_分析"

            _create_ppm_analysis_sheet(
                workbook=workbook,
                analysis_sheet_name=analysis_sheet_name,
                used_sheet_names=info['used_sheet_names'],
                debug_log=debug_log
            )

            _create_composition_ratio_sheet(
                workbook=workbook,
                analysis_sheet_name=analysis_sheet_name,
                used_sheet_names=info['used_sheet_names'],
                debug_log=debug_log
            )

            _create_yoy_growth_sheet(
                workbook=workbook,
                analysis_sheet_name=analysis_sheet_name,
                used_sheet_names=info['used_sheet_names'],
                debug_log=debug_log
            )

            _create_ebitda_sheet(
                workbook=workbook,
                analysis_sheet_name=analysis_sheet_name,
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
    # 1. 「注記」を「連結」に置換
    base_name = sheet_name.replace("注記", "連結")

    # 2. 分析シート名を生成
    analysis_sheet_name = base_name + "_分析"

    # 3. Excelのシート名制限（31文字）をチェック
    if len(analysis_sheet_name) > 31:
        # 31文字を超える場合は、置換後の名前をカットして "_分析" を結合
        # (31 - 3 = 28文字分を本体から取得)
        analysis_sheet_name = base_name[:28] + "_分析"

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

    # 1. 有効な年度の特定と、勘定科目データのアグリゲーション（同じラベルの項目をマージ）
    valid_periods = set()
    label_to_data = {}   # display_label -> {period -> {dim -> value}}
    label_info = {}      # display_label -> {full_path, depth, pref_label, el}
    label_order = []     # 登場順を保持

    for full_path_data in ordered_keys:
        full_path, pref_label = full_path_data
        el = full_path.split('::')[-1]
        if '|' in el:
            el = el.split('|')[0]

        # 不要な要素タイプをスキップ
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

        # ラベルから不要な接尾辞を削除
        display_label = label
        display_label = display_label.replace(' [目次項目]', '').replace(' [タイトル項目]', '')
        display_label = display_label.replace('（IFRS）', '').replace('(IFRS)', '')
        display_label = display_label.replace('、経営指標等', '')
        display_label = display_label.replace('、流動資産', '').replace('、非流動資産', '')
        display_label = display_label.replace('、流動負債', '').replace('、非流動負債', '')
        display_label = display_label.strip()

        depth = len(full_path.split('::')) - 1
        
        if display_label not in label_to_data:
            label_to_data[display_label] = {}
            label_order.append(display_label)
        
        # メタデータを更新（後から出現したものを優先）
        label_info[display_label] = {
            'full_path': full_path,
            'depth': depth,
            'pref_label': pref_label,
            'el': el
        }

        is_sales_item = (display_label == "外部顧客への売上高") or ("売上高" in display_label and "計" in display_label)

        for period in unique_periods:
            for dim in unique_dims:
                found_v = None
                stds_to_check = [current_standard] if current_standard != 'JP_ALL' else ['IFRS', 'JP', 'US', 'JMIS']
                for s in stds_to_check:
                    v = all_years_data[role][full_path].get((s, dim, period))
                    if v is not None:
                        found_v = v
                        break
                
                if found_v is not None and found_v != "":
                    if period not in label_to_data[display_label]:
                        label_to_data[display_label][period] = {}
                    
                    # データをマージ（後から出現したパスのデータで上書き）
                    label_to_data[display_label][period][dim] = found_v
                    
                    # 数値チェックと売上高判定による有効年度追加
                    val_clean = unicodedata.normalize('NFKC', str(found_v)).replace(',', '').strip()
                    try:
                        if val_clean and not any(c.isalpha() for c in val_clean):
                            if is_sales_item:
                                valid_periods.add(period)
                    except:
                        pass

    # 全期間通して一度もデータ（数値・非数値問わず）がない項目を削除
    final_label_order = [l for l in label_order if label_to_data[l]]

    # 有効年度を確定（売上高がない場合は全年度を使用）
    sorted_valid_periods = sorted(list(valid_periods))
    if not sorted_valid_periods:
        sorted_valid_periods = unique_periods

    debug_log(f"[Segment Analysis] Aligned to {len(sorted_valid_periods)} valid periods by merging duplicate labels")

    # 3. データ行を作成
    seen_rows_analysis = set()
    for d_label in final_label_order:
        info = label_info[d_label]
        it_depth = info['depth']
        indent_prefix = "　" * it_depth
        it_el = info['el']
        it_pref_label = info['pref_label']
        it_full_path = info['full_path'] # CF判定用
        it_fp_data = (it_full_path, it_pref_label)

        for period in sorted_valid_periods:
            row_data_analysis = [indent_prefix + d_label, period]
            
            # マージ済みのデータを取得
            period_data = label_to_data[d_label].get(period, {})

            for dim in unique_dims:
                val = period_data.get(dim, "")

                if val:
                    val_clean = unicodedata.normalize('NFKC', str(val)).replace(',', '').strip()
                    try:
                        if val_clean and not any(c.isalpha() for c in val_clean):
                            val = float(val_clean)
                    except:
                        pass
                
                row_data_analysis.append(val)

            # 有効年度内であれば、データが空でも行を出力（ただし項目自体が全期間空の場合はスルー済み）
            # 重複チェックはラベルと年度の組み合わせで行う
            row_key = (d_label, period)
            if row_key in seen_rows_analysis:
                continue
            seen_rows_analysis.add(row_key)
            aws.append(row_data_analysis)

        # キャッシュ・フロー計算書の場合、特定の要素で停止
        if 'キャッシュ・フロー' in sheet_name and 'CashAndCashEquivalents' in it_el:
            if it_pref_label and it_pref_label.endswith(('periodEndLabel', 'totalLabel')):
                # Check if this is at natural end of hierarchy
                try:
                    current_idx = ordered_keys.index(it_fp_data)
                    if current_idx >= len(ordered_keys) - 1:
                        break
                except ValueError:
                    # Should not happen if it's in final_label_order
                    pass

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

    売上（計）とセグメント利益（利益/損失を含む最初のラベル）を
    分析シートから動的に検出し、直近11年分（最新年度 - 10年 ～ 最新年度）の
    データを参照するPPM分析用シートを生成する。

    Args:
        workbook: openpyxlワークブック
        analysis_sheet_name: 参照元の分析シート名
        used_sheet_names: 使用済みシート名のセット
        debug_log: デバッグログ関数
    """
    import re
    import math
    import datetime
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BubbleChart, Reference, Series

    # --- PPM分析シート名を生成 ---
    ppm_sheet_name = analysis_sheet_name + "_PPM分析用"
    if len(ppm_sheet_name) > 31:
        ppm_sheet_name = analysis_sheet_name[:18] + "_PPM分析用"

    debug_log(f"[PPM Analysis] Creating PPM analysis sheet: {ppm_sheet_name}")

    # Check if analysis sheet exists
    if analysis_sheet_name not in workbook.sheetnames:
        debug_log(f"[PPM Analysis] Analysis sheet '{analysis_sheet_name}' not found, skipping PPM sheet")
        return

    analysis_ws = workbook[analysis_sheet_name]
    ppm_ws = workbook.create_sheet(title=ppm_sheet_name)
    used_sheet_names.add(ppm_sheet_name)

    escaped_sheet_name = analysis_sheet_name.replace("'", "''")
    max_col = analysis_ws.max_column

    # -----------------------------------------------------------------------
    # 1. analysis_ws を走査して (ラベル, 年度) -> 行番号 のルックアップを構築
    # -----------------------------------------------------------------------
    unique_labels_ordered = []   # 上から順に初登場のラベル
    lookup = {}                  # (normalized_label, year_int) -> row_index
    max_year = -1

    def _extract_year(period_val):
        """period セルの値から年（int）を取り出す。取得できなければ None。"""
        if period_val is None:
            return None
        if hasattr(period_val, 'year'):
            return period_val.year
        s = str(period_val)
        if '-' in s:
            try:
                return int(s.split('-')[0])
            except ValueError:
                pass
        m = re.search(r'(\d{4})', s)
        return int(m.group(1)) if m else None

    for r in range(2, analysis_ws.max_row + 1):
        label_val = analysis_ws.cell(r, 1).value
        period_val = analysis_ws.cell(r, 2).value
        if not label_val:
            continue
        norm_label = str(label_val).strip()

        # 初登場ラベルだけ順序リストに追加
        if not unique_labels_ordered or unique_labels_ordered[-1] != norm_label:
            unique_labels_ordered.append(norm_label)

        year = _extract_year(period_val)
        if year:
            lookup[(norm_label, year)] = r
            if year > max_year:
                max_year = year

    if max_year == -1:
        debug_log("[PPM Analysis] No years found in analysis sheet, skipping PPM sheet")
        return

    # -----------------------------------------------------------------------
    # 2. 対象ラベルを先頭から順に検索
    #    売上  : 最初に見つかる「計」
    #    利益  : 最初に見つかる「利益」または「損失」を含むラベル
    # -----------------------------------------------------------------------
    target_sales_label = None
    target_profit_label = None
    for label in unique_labels_ordered:
        if target_sales_label is None and label == "計":
            target_sales_label = label
        if target_profit_label is None and ("利益" in label or "損失" in label):
            target_profit_label = label
        if target_sales_label and target_profit_label:
            break

    debug_log(f"[PPM Analysis] max_year={max_year}, Sales label='{target_sales_label}', Profit label='{target_profit_label}'")

    # -----------------------------------------------------------------------
    # 3. 11年分の年度リスト（昇順: max_year-10 ～ max_year）
    # -----------------------------------------------------------------------
    NUM_YEARS = 11
    target_years = list(range(max_year - 10, max_year + 1))

    # 各年の analysis_ws 行番号（データなし年は None）
    sales_src_rows  = [lookup.get((target_sales_label,  y)) if target_sales_label  else None for y in target_years]
    profit_src_rows = [lookup.get((target_profit_label, y)) if target_profit_label else None for y in target_years]

    # -----------------------------------------------------------------------
    # 4. ppm_ws の構築
    # -----------------------------------------------------------------------

    # --- ヘッダー行 ---
    header_row = []
    for col_idx in range(1, max_col + 1):
        cl = get_column_letter(col_idx)
        header_row.append(f"=IF('{escaped_sheet_name}'!{cl}1=\"\",\"\",'{escaped_sheet_name}'!{cl}1)")
    ppm_ws.append(header_row)

    # --- 売上行 (11行) ---
    sales_start_row = ppm_ws.max_row + 1
    for idx, src_row in enumerate(sales_src_rows):
        data_row = ["　売上"]
        for col_idx in range(2, max_col + 1):
            cl = get_column_letter(col_idx)
            if col_idx == 2:  # 年度列
                formula = f"='{escaped_sheet_name}'!B{src_row}" if src_row else target_years[idx]
            else:
                formula = (f"=IF('{escaped_sheet_name}'!{cl}{src_row}=\"\",\"\",'{escaped_sheet_name}'!{cl}{src_row})"
                           if src_row else "")
            data_row.append(formula)
        ppm_ws.append(data_row)
    sales_end_row = ppm_ws.max_row

    # --- セグメント利益行 (11行) ---
    profit_start_row = ppm_ws.max_row + 1
    for idx, src_row in enumerate(profit_src_rows):
        data_row = ["　セグメント利益"]
        for col_idx in range(2, max_col + 1):
            cl = get_column_letter(col_idx)
            if col_idx == 2:  # 年度列
                formula = f"='{escaped_sheet_name}'!B{src_row}" if src_row else target_years[idx]
            else:
                formula = (f"=IF('{escaped_sheet_name}'!{cl}{src_row}=\"\",\"\",'{escaped_sheet_name}'!{cl}{src_row})"
                           if src_row else "")
            data_row.append(formula)
        ppm_ws.append(data_row)
    profit_end_row = ppm_ws.max_row

    # --- 空行区切り ---
    ppm_ws.append([""] * max_col)

    # -----------------------------------------------------------------------
    # 5. 売上高対前年増加率 (11行)
    # -----------------------------------------------------------------------
    growth_start_row = ppm_ws.max_row + 1
    growth_rates = []   # growth_rates[idx][col_idx] = float or None

    for idx in range(NUM_YEARS):
        year_row_ref = sales_start_row + idx
        growth_row = ["売上高対前年増加率"]
        year_growth_rates = [None, None]   # col 0(A) と col 1(B) はラベル/年度

        for col_idx in range(2, max_col + 1):
            cl = get_column_letter(col_idx)
            if col_idx == 2:
                formula = f"=B{year_row_ref}"
                year_growth_rates.append(None)
            elif idx == 0:
                formula = ""
                year_growth_rates.append(None)
            else:
                cur_ref  = sales_start_row + idx
                prev_ref = sales_start_row + idx - 1
                formula  = (f"=IF(OR({cl}{cur_ref}=\"\",{cl}{prev_ref}=\"\"),"
                            f"\"\",{cl}{cur_ref}/{cl}{prev_ref}-1)")
                # 実値（軸計算用）
                cur_src  = sales_src_rows[idx]
                prev_src = sales_src_rows[idx - 1]
                cur_val  = analysis_ws.cell(cur_src,  col_idx).value if cur_src  else None
                prev_val = analysis_ws.cell(prev_src, col_idx).value if prev_src else None
                if isinstance(cur_val, (int, float)) and isinstance(prev_val, (int, float)) and prev_val != 0:
                    year_growth_rates.append(cur_val / prev_val - 1)
                else:
                    year_growth_rates.append(None)
            growth_row.append(formula)

        growth_rates.append(year_growth_rates)
        ppm_ws.append(growth_row)

    growth_end_row = ppm_ws.max_row

    # --- 空行区切り ---
    ppm_ws.append([""] * max_col)

    # -----------------------------------------------------------------------
    # 6. 売上高利益率 (11行)
    # -----------------------------------------------------------------------
    margin_start_row = ppm_ws.max_row + 1
    profit_margins = []   # profit_margins[idx][col_idx] = float or None

    for idx in range(NUM_YEARS):
        year_row_ref = sales_start_row + idx
        margin_row = ["売上高利益率"]
        year_profit_margins = [None, None]   # col 0(A) と col 1(B)

        for col_idx in range(2, max_col + 1):
            cl = get_column_letter(col_idx)
            if col_idx == 2:
                formula = f"=B{year_row_ref}"
                year_profit_margins.append(None)
            else:
                s_ref = sales_start_row  + idx
                p_ref = profit_start_row + idx
                formula = (f"=IF(OR({cl}{p_ref}=\"\",{cl}{s_ref}=\"\"),"
                           f"\"\",{cl}{p_ref}/{cl}{s_ref})")
                # 実値（軸計算用）
                s_src = sales_src_rows[idx]
                p_src = profit_src_rows[idx]
                s_val = analysis_ws.cell(s_src, col_idx).value if s_src else None
                p_val = analysis_ws.cell(p_src, col_idx).value if p_src else None
                if isinstance(s_val, (int, float)) and isinstance(p_val, (int, float)) and s_val != 0:
                    year_profit_margins.append(p_val / s_val)
                else:
                    year_profit_margins.append(None)
            margin_row.append(formula)

        profit_margins.append(year_profit_margins)
        ppm_ws.append(margin_row)

    margin_end_row = ppm_ws.max_row

    # 売上実値（セグメント完全性チェック用）
    sales_values = []
    for idx in range(NUM_YEARS):
        year_sales = [None, None]
        for col_idx in range(2, max_col + 1):
            if col_idx == 2:
                year_sales.append(None)
            else:
                s_src = sales_src_rows[idx]
                s_val = analysis_ws.cell(s_src, col_idx).value if s_src else None
                year_sales.append(s_val if isinstance(s_val, (int, float)) else None)
        sales_values.append(year_sales)

    # -----------------------------------------------------------------------
    # 7. 書式設定
    # -----------------------------------------------------------------------
    ppm_ws.freeze_panes = 'B2'
    ppm_ws.column_dimensions['A'].width = 31
    for col_idx in range(2, max_col + 1):
        ppm_ws.column_dimensions[get_column_letter(col_idx)].width = 12

    for row_idx in range(sales_start_row, profit_end_row + 1):
        for col_idx in range(3, max_col + 1):
            ppm_ws.cell(row_idx, col_idx).number_format = r'#,##0_ ;[Red]\-#,##0 '

    for row_idx in list(range(growth_start_row, growth_end_row + 1)) + list(range(margin_start_row, margin_end_row + 1)):
        for col_idx in range(3, max_col + 1):
            ppm_ws.cell(row_idx, col_idx).number_format = '0%'

    # -----------------------------------------------------------------------
    # 8. チャート用集約データエリア
    # -----------------------------------------------------------------------
    ppm_ws.append([""] * max_col)
    ppm_ws.append([""] * max_col)
    data_start_row = ppm_ws.max_row + 1

    LATEST_IDX     = NUM_YEARS - 1   # 最新年（インデックス10）
    FIVE_AGO_IDX   = NUM_YEARS - 6   # 5年前（インデックス5）
    FIVE_AGO_OFFSET = 5

    # 「報告セグメント」列を chart_end_col に設定
    chart_end_col = max_col
    for col_idx in range(3, max_col + 1):
        hv = analysis_ws.cell(1, col_idx).value
        if hv and "報告セグメント" in str(hv) and "以外" not in str(hv):
            chart_end_col = col_idx
            break

    # チャートタイトル用の期間値
    latest_src = sales_src_rows[LATEST_IDX] or profit_src_rows[LATEST_IDX]
    year_val_for_title = analysis_ws.cell(latest_src, 2).value if latest_src else target_years[LATEST_IDX]

    def _fmt_year_str(yv):
        if yv is None:
            return ""
        if hasattr(yv, 'strftime'):
            return yv.strftime('%Y/%m')
        s = str(yv)
        if '-' in s:
            try:
                return datetime.datetime.strptime(s, '%Y-%m-%d').strftime('%Y/%m')
            except ValueError:
                return s[:7].replace('-', '/')
        return s[:4]

    def _valid_cols(arr, year_idx):
        """3指標すべてが揃っている列のリストを返す"""
        result = []
        for ci in range(3, chart_end_col + 1):
            if (ci < len(profit_margins[year_idx])  and profit_margins[year_idx][ci]  is not None and
                ci < len(growth_rates[year_idx])    and growth_rates[year_idx][ci]    is not None and
                ci < len(sales_values[year_idx])    and sales_values[year_idx][ci]    is not None):
                result.append(ci)
        return result

    def _append_data_section(year_idx, sales_row, profit_row, growth_row, margin_row):
        """集約データ4行（ヘッダ・利益率・成長率・売上）を追加して開始行を返す"""
        sec_start = ppm_ws.max_row + 1
        vcols = _valid_cols(profit_margins if year_idx == LATEST_IDX else profit_margins, year_idx)

        # ヘッダ行
        hrow = [f"=B{sales_row}", "セグメント名"]
        for ci in vcols:
            hrow.append(f"={get_column_letter(ci)}1")
        if vcols and vcols[-1] == chart_end_col:
            hrow[-1] = "計"
        ppm_ws.append(hrow)

        # 利益率行
        mrow = [f"=B{margin_row}", f"=A{margin_row}"]
        for ci in vcols:
            mrow.append(f"={get_column_letter(ci)}{margin_row}")
        ppm_ws.append(mrow)

        # 成長率行
        grow = [f"=B{growth_row}", f"=A{growth_row}"]
        for ci in vcols:
            grow.append(f"={get_column_letter(ci)}{growth_row}")
        ppm_ws.append(grow)

        # 売上行
        srow = [f"=B{sales_row}", f"=TRIM(A{sales_row})"]
        for ci in vcols:
            cl = get_column_letter(ci)
            srow.append(f"={cl}{sales_row}*1%" if ci == chart_end_col else f"={cl}{sales_row}")
        ppm_ws.append(srow)

        sec_end = ppm_ws.max_row
        sec_max_col = max(3, 2 + len(vcols))

        # 書式
        for ci in range(3, sec_max_col + 1):
            cl = get_column_letter(ci)
            ppm_ws[f'{cl}{sec_start + 1}'].number_format = '0%'
            ppm_ws[f'{cl}{sec_start + 2}'].number_format = '0%'
            ppm_ws[f'{cl}{sec_start + 3}'].number_format = r'#,##0_);[Red](#,##0)'

        return sec_start, sec_end, sec_max_col, vcols

    # 最新年データ
    latest_sales_row  = sales_end_row
    latest_profit_row = profit_end_row
    latest_growth_row = growth_end_row
    latest_margin_row = margin_end_row

    lat_start, lat_end, lat_max_col, _ = _append_data_section(
        LATEST_IDX, latest_sales_row, latest_profit_row, latest_growth_row, latest_margin_row
    )

    # -----------------------------------------------------------------------
    # 9. 軸範囲計算（最新年 + 5年前 の共通スケール）
    # -----------------------------------------------------------------------
    def _axis_values(arr, year_indices):
        vals = []
        for yi in year_indices:
            if 0 <= yi < len(arr):
                for v in arr[yi]:
                    if v is not None:
                        vals.append(v)
        return vals

    def _rounded_range(vals):
        if not vals:
            return -0.05, 0.40
        mn, mx = min(vals), max(vals)
        return math.floor(mn / 0.05) * 0.05, math.ceil(mx / 0.05) * 0.05

    rel_years = [LATEST_IDX]
    if NUM_YEARS > FIVE_AGO_OFFSET:
        rel_years.append(FIVE_AGO_IDX)

    common_x_min, common_x_max = _rounded_range(_axis_values(profit_margins, rel_years))
    common_y_min, common_y_max = _rounded_range(_axis_values(growth_rates,  rel_years))

    # -----------------------------------------------------------------------
    # 10. バブルチャートを作成するヘルパー
    # -----------------------------------------------------------------------
    def _make_chart(title_str):
        chart = BubbleChart()
        chart.style = 2
        chart.height, chart.width = 15, 15
        chart.title = f"PPM分析 {title_str}"
        chart.x_axis.title = "売上高利益率"
        chart.y_axis.title = "売上高対前年増加率"
        chart.x_axis.tickLblPos = "nextTo"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.scaling.min = common_x_min
        chart.x_axis.scaling.max = common_x_max
        chart.y_axis.scaling.min = common_y_min
        chart.y_axis.scaling.max = common_y_max
        chart.legend = None
        return chart

    def _add_series(chart, ws, sec_start, sec_max_col):
        xv = Reference(ws, min_col=3, min_row=sec_start + 1, max_col=sec_max_col, max_row=sec_start + 1)
        yv = Reference(ws, min_col=3, min_row=sec_start + 2, max_col=sec_max_col, max_row=sec_start + 2)
        sz = Reference(ws, min_col=3, min_row=sec_start + 3, max_col=sec_max_col, max_row=sec_start + 3)
        chart.series.append(Series(values=yv, xvalues=xv, zvalues=sz, title=""))

    chart_latest = _make_chart(_fmt_year_str(year_val_for_title))
    _add_series(chart_latest, ppm_ws, lat_start, lat_max_col)

    # -----------------------------------------------------------------------
    # 11. 5年前データセクション（存在する場合）
    # -----------------------------------------------------------------------
    ppm_ws.append([""] * max_col)

    chart_5y = None
    if NUM_YEARS > FIVE_AGO_OFFSET:
        five_sales_row  = sales_end_row  - FIVE_AGO_OFFSET
        five_profit_row = profit_end_row - FIVE_AGO_OFFSET
        five_growth_row = growth_end_row - FIVE_AGO_OFFSET
        five_margin_row = margin_end_row - FIVE_AGO_OFFSET

        y_val_5y = (analysis_ws.cell(sales_src_rows[FIVE_AGO_IDX], 2).value
                    if sales_src_rows[FIVE_AGO_IDX] else target_years[FIVE_AGO_IDX])

        five_start, five_end, five_max_col, _ = _append_data_section(
            FIVE_AGO_IDX, five_sales_row, five_profit_row, five_growth_row, five_margin_row
        )

        chart_5y = _make_chart(_fmt_year_str(y_val_5y))
        _add_series(chart_5y, ppm_ws, five_start, five_max_col)
        debug_log("[PPM Analysis] Added 5-year comparison section")

    # -----------------------------------------------------------------------
    # 12. チャートをシートに配置
    # -----------------------------------------------------------------------
    chart_row = ppm_ws.max_row + 2
    ppm_ws.add_chart(chart_latest, f'B{chart_row}')
    if chart_5y:
        ppm_ws.add_chart(chart_5y, f'I{chart_row}')

    debug_log(f"[PPM Analysis] Completed PPM analysis sheet: {ppm_sheet_name}")


def _create_composition_ratio_sheet(workbook, analysis_sheet_name, used_sheet_names, debug_log):
    """
    構成比シートを作成（内部関数）

    分析シートの各セルを「報告セグメント」列で割った構成比を表示するシートを生成する。
    数式: =IF(OR('分析'!C2="", '分析'!$H2=""), "", '分析'!C2/'分析'!$H2)

    Args:
        workbook: openpyxlワークブック
        analysis_sheet_name: 参照元の分析シート名
        used_sheet_names: 使用済みシート名のセット
        debug_log: デバッグログ関数
    """
    from openpyxl.utils import get_column_letter

    comp_sheet_name = analysis_sheet_name + "_構成比"
    if len(comp_sheet_name) > 31:
        comp_sheet_name = analysis_sheet_name[:25] + "_構成比"

    debug_log(f"[Composition] Creating composition ratio sheet: {comp_sheet_name}")

    if analysis_sheet_name not in workbook.sheetnames:
        debug_log(f"[Composition] Analysis sheet '{analysis_sheet_name}' not found, skipping")
        return

    analysis_ws = workbook[analysis_sheet_name]
    comp_ws = workbook.create_sheet(title=comp_sheet_name)
    used_sheet_names.add(comp_sheet_name)

    escaped = analysis_sheet_name.replace("'", "''")
    max_col = analysis_ws.max_column
    max_row = analysis_ws.max_row

    # 「報告セグメント」列を動的に検出
    denom_col = max_col
    for col_idx in range(3, max_col + 1):
        hv = analysis_ws.cell(1, col_idx).value
        if hv and "報告セグメント" in str(hv) and "以外" not in str(hv):
            denom_col = col_idx
            break

    denom_cl = get_column_letter(denom_col)

    # ヘッダ行（行1）: 勘定科目・年度・各セグメント列をそのままコピー参照（報告セグメントまで）
    header = ["勘定科目", "年度"]
    for col_idx in range(3, denom_col + 1):
        cl = get_column_letter(col_idx)
        header.append(f"=IF('{escaped}'!{cl}1=\"\",\"\",'{escaped}'!{cl}1)")
    comp_ws.append(header)

    # データ行（行2以降）
    for row in range(2, max_row + 1):
        row_data = []
        # A列: 勘定科目
        row_data.append(f"='{escaped}'!A{row}")
        # B列: 年度
        row_data.append(f"='{escaped}'!B{row}")
        # C列以降: 構成比（報告セグメント列まで）
        for col_idx in range(3, denom_col + 1):
            cl = get_column_letter(col_idx)
            formula = (
                f"=IF(OR('{escaped}'!{cl}{row}=\"\","
                f"'{escaped}'!${denom_cl}{row}=\"\"),"
                f"\"\","
                f"'{escaped}'!{cl}{row}/'{escaped}'!${denom_cl}{row})"
            )
            row_data.append(formula)
        comp_ws.append(row_data)

    # 書式設定
    comp_ws.freeze_panes = 'B2'
    comp_ws.column_dimensions['A'].width = 31
    for col_idx in range(2, denom_col + 1):
        comp_ws.column_dimensions[get_column_letter(col_idx)].width = 12

    # 数値セル（C列以降、2行目以降）に % 書式を設定
    for row in comp_ws.iter_rows(min_row=2, min_col=3, max_col=denom_col):
        for cell in row:
            cell.number_format = '0%'

    debug_log(f"[Composition] Completed composition ratio sheet: {comp_sheet_name}")


def _create_yoy_growth_sheet(workbook, analysis_sheet_name, used_sheet_names, debug_log):
    """
    対前年増加率シートを作成（内部関数）

    分析シートの各セルについて、当年/前年 - 1 の対前年増加率を表示するシートを生成する。
    数式: =IF(OR('分析'!C2="", '分析'!C3=""), "", '分析'!C2/'分析'!C3-1)
    全カラム（max_col まで）を出力する。

    Args:
        workbook: openpyxlワークブック
        analysis_sheet_name: 参照元の分析シート名
        used_sheet_names: 使用済みシート名のセット
        debug_log: デバッグログ関数
    """
    from openpyxl.utils import get_column_letter

    yoy_sheet_name = analysis_sheet_name + "_対前年増加率"
    if len(yoy_sheet_name) > 31:
        yoy_sheet_name = analysis_sheet_name[:22] + "_対前年増加率"

    debug_log(f"[YoY] Creating YoY growth sheet: {yoy_sheet_name}")

    if analysis_sheet_name not in workbook.sheetnames:
        debug_log(f"[YoY] Analysis sheet '{analysis_sheet_name}' not found, skipping")
        return

    analysis_ws = workbook[analysis_sheet_name]
    yoy_ws = workbook.create_sheet(title=yoy_sheet_name)
    used_sheet_names.add(yoy_sheet_name)

    escaped = analysis_sheet_name.replace("'", "''")
    max_col = analysis_ws.max_column
    max_row = analysis_ws.max_row

    # 各勘定科目グループの先頭行（一番古い年度）を検出
    # A列の値が変わる行 = 新しい勘定科目の先頭
    first_rows_of_group = set()
    prev_label = object()  # sentinel
    for row in range(2, max_row + 1):
        label = analysis_ws.cell(row, 1).value
        if label != prev_label:
            first_rows_of_group.add(row)
            prev_label = label

    # ヘッダ行（行1）
    header = ["勘定科目", "年度"]
    for col_idx in range(3, max_col + 1):
        cl = get_column_letter(col_idx)
        header.append(f"=IF('{escaped}'!{cl}1=\"\",\"\",'{escaped}'!{cl}1)")
    yoy_ws.append(header)

    # 行2以降: 各勘定の先頭行はデータ列を空白、それ以外は対前年増加率
    for row in range(2, max_row + 1):
        row_data = [f"='{escaped}'!A{row}", f"='{escaped}'!B{row}"]
        if row in first_rows_of_group:
            # 一番古い年度: 前年データなしのため空白
            row_data += [""] * (max_col - 2)
        else:
            for col_idx in range(3, max_col + 1):
                cl = get_column_letter(col_idx)
                formula = (
                    f"=IF(OR('{escaped}'!{cl}{row - 1}=\"\","
                    f"'{escaped}'!{cl}{row}=\"\"),"
                    f"\"\","
                    f"'{escaped}'!{cl}{row - 1}/'{escaped}'!{cl}{row}-1)"
                )
                row_data.append(formula)
        yoy_ws.append(row_data)

    # 書式設定
    yoy_ws.freeze_panes = 'B2'
    yoy_ws.column_dimensions['A'].width = 31
    for col_idx in range(2, max_col + 1):
        yoy_ws.column_dimensions[get_column_letter(col_idx)].width = 12

    # 数値セル（C列以降、3行目以降）に % 書式を設定
    for row in yoy_ws.iter_rows(min_row=3, min_col=3, max_col=max_col):
        for cell in row:
            cell.number_format = '0%'

    debug_log(f"[YoY] Completed YoY growth sheet: {yoy_sheet_name}")


def _create_ebitda_sheet(workbook, analysis_sheet_name, used_sheet_names, debug_log):
    """
    EBITDA分析シートを作成（内部関数）

    分析シートから売上高（「計」）・セグメント利益（最初の「利益」or「損失」）・
    償却費/償却額を含む全勘定・減損損失を動的に検出し、EBITDAを計算するシートを生成する。

    Args:
        workbook: openpyxlワークブック
        analysis_sheet_name: 参照元の分析シート名
        used_sheet_names: 使用済みシート名のセット
        debug_log: デバッグログ関数
    """
    import re
    from openpyxl.utils import get_column_letter

    ebitda_sheet_name = analysis_sheet_name + "_EBITDA"
    if len(ebitda_sheet_name) > 31:
        ebitda_sheet_name = analysis_sheet_name[:24] + "_EBITDA"

    debug_log(f"[EBITDA] Creating EBITDA sheet: {ebitda_sheet_name}")

    if analysis_sheet_name not in workbook.sheetnames:
        debug_log(f"[EBITDA] Analysis sheet '{analysis_sheet_name}' not found, skipping")
        return

    analysis_ws = workbook[analysis_sheet_name]
    ebitda_ws = workbook.create_sheet(title=ebitda_sheet_name)
    used_sheet_names.add(ebitda_sheet_name)

    escaped = analysis_sheet_name.replace("'", "''")
    max_col = analysis_ws.max_column

    # -----------------------------------------------------------------------
    # 1. analysis_ws を走査してラベルと年度のルックアップを構築
    # -----------------------------------------------------------------------
    unique_labels_ordered = []
    lookup = {}
    max_year = -1

    def _extract_year(period_val):
        if period_val is None:
            return None
        if hasattr(period_val, 'year'):
            return period_val.year
        s = str(period_val)
        if '-' in s:
            try:
                return int(s.split('-')[0])
            except ValueError:
                pass
        m = re.search(r'(\d{4})', s)
        return int(m.group(1)) if m else None

    for r in range(2, analysis_ws.max_row + 1):
        label_val = analysis_ws.cell(r, 1).value
        period_val = analysis_ws.cell(r, 2).value
        if not label_val:
            continue
        norm_label = str(label_val).strip()
        if not unique_labels_ordered or unique_labels_ordered[-1] != norm_label:
            unique_labels_ordered.append(norm_label)
        year = _extract_year(period_val)
        if year:
            lookup[(norm_label, year)] = r
            if year > max_year:
                max_year = year

    if max_year == -1:
        debug_log("[EBITDA] No years found in analysis sheet, skipping")
        return

    # -----------------------------------------------------------------------
    # 2. 対象ラベルを検索
    # -----------------------------------------------------------------------
    target_sales_label = None
    target_profit_label = None
    amortization_labels = []
    impairment_label = None

    for label in unique_labels_ordered:
        if target_sales_label is None and label == "計":
            target_sales_label = label
        if target_profit_label is None and ("利益" in label or "損失" in label):
            target_profit_label = label
        if ("償却費" in label or "償却額" in label) and label not in amortization_labels:
            amortization_labels.append(label)
        if impairment_label is None and "減損損失" in label:
            impairment_label = label

    # EBITDA構成要素: セグメント利益 + 償却費/償却額 + 減損損失
    ebitda_items = []
    if target_profit_label:
        ebitda_items.append(target_profit_label)
    ebitda_items.extend(amortization_labels)
    if impairment_label and impairment_label not in ebitda_items:
        ebitda_items.append(impairment_label)

    debug_log(f"[EBITDA] max_year={max_year}, Sales='{target_sales_label}', "
              f"EBITDA items={ebitda_items}")

    # -----------------------------------------------------------------------
    # 3. 11年分の年度リスト（昇順）
    # -----------------------------------------------------------------------
    NUM_YEARS = 11
    target_years = list(range(max_year - 10, max_year + 1))

    # -----------------------------------------------------------------------
    # 4. シートを構築
    # -----------------------------------------------------------------------

    # --- ヘッダー行 ---
    header_row = []
    for col_idx in range(1, max_col + 1):
        cl = get_column_letter(col_idx)
        header_row.append(f"=IF('{escaped}'!{cl}1=\"\",\"\",'{escaped}'!{cl}1)")
    ebitda_ws.append(header_row)

    def _write_data_block(label, src_rows_list, display_label=None):
        """11年分データブロックを書き込み、ブロック開始行を返す"""
        block_start = ebitda_ws.max_row + 1
        lbl = display_label if display_label is not None else label
        for idx, src_row in enumerate(src_rows_list):
            row_data = [lbl]
            for col_idx in range(2, max_col + 1):
                cl = get_column_letter(col_idx)
                if col_idx == 2:
                    formula = f"='{escaped}'!B{src_row}" if src_row else ""
                else:
                    formula = (
                        f"=IF('{escaped}'!{cl}{src_row}=\"\",\"\",'{escaped}'!{cl}{src_row})"
                        if src_row else ""
                    )
                row_data.append(formula)
            ebitda_ws.append(row_data)
        return block_start

    # 売上高ブロック
    sales_src_rows = [
        lookup.get((target_sales_label, y)) if target_sales_label else None
        for y in target_years
    ]
    sales_block_start = _write_data_block(
        target_sales_label or "売上高", sales_src_rows, "　売上高"
    )

    # EBITDAアイテムのブロック
    item_block_starts = []
    item_is_negative = []  # 負ののれんを含む場合はTrue
    for item_label in ebitda_items:
        item_src_rows = [lookup.get((item_label, y)) for y in target_years]
        block_start = _write_data_block(item_label, item_src_rows)
        item_block_starts.append(block_start)
        item_is_negative.append("負ののれん" in item_label)

    # --- 空行 ---
    ebitda_ws.append([""] * max_col)

    # --- EBITDAブロック ---
    ebitda_block_start = ebitda_ws.max_row + 1
    for idx in range(NUM_YEARS):
        row_data = ["　EBITDA"]
        row_data.append(f"=B{sales_block_start + idx}")
        for col_idx in range(3, max_col + 1):
            cl = get_column_letter(col_idx)
            if item_block_starts:
                refs = ",".join(
                    f"{cl}{start + idx}"
                    for start, is_neg in zip(item_block_starts, item_is_negative)
                    if not is_neg
                )
                refs_minus = ",".join(
                    f"{cl}{start + idx}"
                    for start, is_neg in zip(item_block_starts, item_is_negative)
                    if is_neg
                )
                if not refs_minus:
                    formula = f"=IF(COUNT({refs})=0,\"\",SUM({refs}))" if refs else ""
                else:
                    count_refs = refs if refs else refs_minus
                    sum_part = f"SUM({refs})-SUM({refs_minus})" if refs else f"-SUM({refs_minus})"
                    formula = f"=IF(COUNT({count_refs})=0,\"\",{sum_part})"
            else:
                formula = ""
            row_data.append(formula)
        ebitda_ws.append(row_data)
    ebitda_block_end = ebitda_ws.max_row

    # --- 空行 ---
    ebitda_ws.append([""] * max_col)

    # --- EBITDA/売上高ブロック ---
    ratio_block_start = ebitda_ws.max_row + 1
    for idx in range(NUM_YEARS):
        row_data = ["EBITDA/売上高"]
        row_data.append(f"=B{sales_block_start + idx}")
        for col_idx in range(3, max_col + 1):
            cl = get_column_letter(col_idx)
            ebitda_row = ebitda_block_start + idx
            sales_row = sales_block_start + idx
            formula = (
                f"=IF(OR({cl}{ebitda_row}=\"\",{cl}{sales_row}=\"\"),"
                f"\"\",{cl}{ebitda_row}/{cl}{sales_row})"
            )
            row_data.append(formula)
        ebitda_ws.append(row_data)

    # --- 空行 ---
    ebitda_ws.append([""] * max_col)

    # --- EBITDA対前年増加率ブロック ---
    for idx in range(NUM_YEARS):
        row_data = ["EBITDA対前年増加率"]
        row_data.append(f"=B{sales_block_start + idx}")
        if idx == 0:
            # 最古年度は前年データなしのため空白
            row_data += [""] * (max_col - 2)
        else:
            for col_idx in range(3, max_col + 1):
                cl = get_column_letter(col_idx)
                cur_row = ebitda_block_start + idx
                prev_row = ebitda_block_start + idx - 1
                formula = (
                    f"=IF(OR({cl}{cur_row}=\"\",{cl}{prev_row}=\"\"),"
                    f"\"\",{cl}{cur_row}/{cl}{prev_row}-1)"
                )
                row_data.append(formula)
        ebitda_ws.append(row_data)

    # -----------------------------------------------------------------------
    # 5. 書式設定
    # -----------------------------------------------------------------------
    # 数値書式（ヘッダー以降〜EBITDAブロック末: 金額）
    for row in ebitda_ws.iter_rows(min_row=2, max_row=ebitda_block_end, min_col=3, max_col=max_col):
        for cell in row:
            cell.number_format = r'#,##0_ ;[Red]\-#,##0 '

    # % 書式（EBITDA/売上高、EBITDA対前年増加率）
    for row in ebitda_ws.iter_rows(min_row=ratio_block_start, max_row=ebitda_ws.max_row, min_col=3, max_col=max_col):
        for cell in row:
            cell.number_format = '0%'

    # ウィンドウ枠固定
    ebitda_ws.freeze_panes = 'B2'

    # 列幅
    ebitda_ws.column_dimensions['A'].width = 31
    for col_idx in range(2, max_col + 1):
        ebitda_ws.column_dimensions[get_column_letter(col_idx)].width = 12

    debug_log(f"[EBITDA] Completed EBITDA sheet: {ebitda_sheet_name}")
