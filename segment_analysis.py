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
            global_element_period_values=info.get('global_element_period_values', {}),
            debug_log=debug_log
        )

        # 元の注記シートの末尾にセグメント別研究開発費を追加
        _append_rd_to_notes_sheet(
            workbook=workbook,
            sheet_name=info['sheet_name'],
            sorted_role_cols=info['sorted_role_cols'],
            global_element_period_values=info.get('global_element_period_values', {}),
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

        elif 'IFRS' in info['sheet_name']:
            base_name = info['sheet_name'].replace("注記", "連結")
            analysis_sheet_name = base_name + "_分析"
            if len(analysis_sheet_name) > 31:
                analysis_sheet_name = base_name[:28] + "_分析"

            _create_ppm_analysis_sheet_ifrs(
                workbook=workbook,
                analysis_sheet_name=analysis_sheet_name,
                used_sheet_names=info['used_sheet_names'],
                debug_log=debug_log
            )

            _create_ebitda_sheet_ifrs(
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


def _create_segment_analysis_sheet(workbook, sheet_name, ordered_keys, all_years_data,
                                   role, sorted_role_cols, role_columns, current_standard,
                                   segment_dict, common_dict, labels_map, used_sheet_names,
                                   global_element_period_values, debug_log):
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

    # 「報告セグメント及びその他の合計」がある場合は「報告セグメント」（単独列）を除去する。
    # 単独の「報告セグメント」列はデータが空になることが多く冗長なため。
    # PPM分析用の chart_end_col も自動的に「報告セグメント及びその他の合計」を参照するようになる。
    _HAS_TOTAL_WITH_OTHER = any(
        d == "報告セグメント及びその他の合計" for d in unique_dims
    )
    if _HAS_TOTAL_WITH_OTHER:
        unique_dims = [d for d in unique_dims if d != "報告セグメント"]

    # IFRS: 報告セグメント合計列が存在しない場合は合成列を追加
    synthetic_total_dim = None
    reporting_dims_for_total = []

    if 'IFRS' in sheet_name:
        has_total = any(
            "報告セグメント" in str(d) and "以外" not in str(d)
            for d in unique_dims
        )
        if not has_total:
            ikai_idx = next(
                (i for i, d in enumerate(unique_dims) if "以外" in str(d)),
                None
            )
            if ikai_idx is not None and ikai_idx > 0:
                reporting_dims_for_total = unique_dims[:ikai_idx]
                unique_dims.insert(ikai_idx, "報告セグメント合計")
                synthetic_total_dim = "報告セグメント合計"
            elif ikai_idx is None and unique_dims:
                reporting_dims_for_total = list(unique_dims)
                unique_dims.append("報告セグメント合計")
                synthetic_total_dim = "報告セグメント合計"

    # 報告セグメント合計列のSUM式用: 列位置を事前計算
    # unique_dims 内の synthetic_total_dim の位置 → シート上の列番号 (A=1, B=2, C=3...)
    synthetic_total_col_idx = None   # シート列番号（1-based）
    synthetic_sum_first_col = None   # SUM範囲の先頭列レター
    synthetic_sum_last_col  = None   # SUM範囲の末尾列レター
    if synthetic_total_dim:
        from openpyxl.utils import get_column_letter as _gcl
        tot_pos = unique_dims.index(synthetic_total_dim)   # 0-based in unique_dims
        synthetic_total_col_idx = 3 + tot_pos              # C=3 が unique_dims[0]
        # reporting_dims_for_total は unique_dims の 0 ~ tot_pos-1
        synthetic_sum_first_col = _gcl(3)                  # C
        synthetic_sum_last_col  = _gcl(3 + tot_pos - 1)   # tot_pos 列分の最後
        del _gcl

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
                if dim == synthetic_total_dim:
                    # 報告セグメント合計: 各報告セグメントの合計を計算
                    total_val = 0.0
                    has_any_val = False
                    for rd in reporting_dims_for_total:
                        rv = period_data.get(rd, "")
                        if rv:
                            rv_clean = unicodedata.normalize('NFKC', str(rv)).replace(',', '').strip()
                            try:
                                if rv_clean and not any(c.isalpha() for c in rv_clean):
                                    total_val += float(rv_clean)
                                    has_any_val = True
                            except:
                                pass
                    val = total_val if has_any_val else ""
                else:
                    val = period_data.get(dim, "")

                    if val:
                        val_clean = unicodedata.normalize('NFKC', str(val)).replace(',', '').strip()
                        try:
                            if val_clean and not any(c.isalpha() for c in val_clean):
                                val = float(val_clean)
                        except:
                            pass

                row_data_analysis.append(val)

            # 数値データが1つもない行（テキストのみ）はスキップ
            has_numeric = any(isinstance(v, (int, float)) for v in row_data_analysis[2:])
            if not has_numeric:
                continue

            # 重複チェックはラベルと年度の組み合わせで行う
            row_key = (d_label, period)
            if row_key in seen_rows_analysis:
                continue
            seen_rows_analysis.add(row_key)

            # 報告セグメント合計列をSUM式で上書き
            if synthetic_total_dim and synthetic_sum_first_col and synthetic_sum_last_col:
                next_row = aws.max_row + 1
                row_data_analysis[synthetic_total_col_idx - 1] = (
                    f"=SUM({synthetic_sum_first_col}{next_row}:{synthetic_sum_last_col}{next_row})"
                )

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

    # プレゼンテーションツリーに含まれない追加項目をシート末尾に追加
    # (研究開発費、設備投資額、従業員数など)
    _EXTRA_SEGMENT_ELEMENTS = [
        ('jpcrp_cor_ResearchAndDevelopmentExpensesResearchAndDevelopmentActivities', '研究開発費'),
        ('jpcrp_cor_CapitalExpendituresOverviewOfCapitalExpendituresEtc', '設備投資額'),
        ('jpcrp_cor_NumberOfEmployees', '従業員数'),
        ('jpcrp_cor_AverageNumberOfTemporaryWorkers', '平均臨時雇用人員'),
    ]
    fact_std = current_standard if current_standard not in ('JP_ALL',) else 'JP'

    # 「報告セグメント及びその他の合計」列の位置を特定（構成比の分母として使う）
    _TOTAL_SUFFIXES = ('合計', '全体', '全社', '消去', '調整', '連結財務諸表')
    _reporting_total_dim = next(
        (d for d in unique_dims if '報告セグメント' in str(d) and '以外' not in str(d) and '合計' in str(d)),
        None
    )
    # 合計列を構成する個別セグメント次元（合計・全体・全社・消去・調整を除く）
    _summable_dims = [
        d for d in unique_dims
        if d != synthetic_total_dim and not any(s in str(d) for s in _TOTAL_SUFFIXES)
    ]

    for extra_el, extra_label in _EXTRA_SEGMENT_ELEMENTS:
        extra_vals = global_element_period_values.get(extra_el, {})
        if not extra_vals:
            continue

        # セグメント次元にデータがあるか確認
        has_segment_data = False
        for period in sorted_valid_periods:
            for dim in unique_dims:
                if dim == synthetic_total_dim:
                    continue
                if extra_vals.get((fact_std, dim, period)) is not None:
                    has_segment_data = True
                    break
            if has_segment_data:
                break
        if not has_segment_data:
            continue

        debug_log(f"[Segment Analysis] Appending '{extra_label}' rows")
        for period in sorted_valid_periods:
            row_data = [extra_label, period]
            # 各dim列の値を収集（後で合計列の補完に使う）
            dim_values = {}
            for dim in unique_dims:
                if dim == synthetic_total_dim:
                    total = 0.0
                    has_any = False
                    for rd in reporting_dims_for_total:
                        v = extra_vals.get((fact_std, rd, period))
                        if v is not None:
                            vc = unicodedata.normalize('NFKC', str(v)).replace(',', '').strip()
                            try:
                                total += float(vc)
                                has_any = True
                            except ValueError:
                                pass
                    dim_values[dim] = total if has_any else ""
                else:
                    v = extra_vals.get((fact_std, dim, period))
                    if v is not None:
                        vc = unicodedata.normalize('NFKC', str(v)).replace(',', '').strip()
                        try:
                            dim_values[dim] = float(vc)
                        except ValueError:
                            dim_values[dim] = v
                    else:
                        dim_values[dim] = ""

            # 「報告セグメント及びその他の合計」列が空の場合、個別セグメントの合計で補完
            if _reporting_total_dim and dim_values.get(_reporting_total_dim) == "":
                seg_sum = 0.0
                has_any_seg = False
                for sd in _summable_dims:
                    sv = dim_values.get(sd, "")
                    if isinstance(sv, (int, float)):
                        seg_sum += sv
                        has_any_seg = True
                if has_any_seg:
                    dim_values[_reporting_total_dim] = seg_sum

            for dim in unique_dims:
                row_data.append(dim_values[dim])
            aws.append(row_data)

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
    from openpyxl.utils import get_column_letter
    for col_idx in range(2, aws.max_column + 1):
        aws.column_dimensions[get_column_letter(col_idx)].width = 12

    debug_log(f"[Segment Analysis] Completed analysis sheet: {analysis_sheet_name} with {aws.max_row - 1} data rows")


def _append_rd_to_notes_sheet(workbook, sheet_name, sorted_role_cols,
                              global_element_period_values, debug_log):
    """元の注記シートの末尾にセグメント別に追加項目（研究開発費等）の行を追加する"""
    if sheet_name not in workbook.sheetnames:
        return

    _EXTRA_SEGMENT_ELEMENTS = [
        ('jpcrp_cor_ResearchAndDevelopmentExpensesResearchAndDevelopmentActivities', '研究開発費'),
        ('jpcrp_cor_CapitalExpendituresOverviewOfCapitalExpendituresEtc', '設備投資額'),
        ('jpcrp_cor_NumberOfEmployees', '従業員数'),
        ('jpcrp_cor_AverageNumberOfTemporaryWorkers', '平均臨時雇用人員'),
    ]

    non_segment_dims = ('全体', '単体', '連結', '全社', '連結財務諸表計上額')
    ws = workbook[sheet_name]
    appended = False

    for extra_el, extra_label in _EXTRA_SEGMENT_ELEMENTS:
        extra_vals = global_element_period_values.get(extra_el, {})
        if not extra_vals:
            continue

        has_segment_data = any(
            extra_vals.get(col_key) is not None
            for col_key in sorted_role_cols
            if len(col_key) == 3 and col_key[1] not in non_segment_dims
        )
        if not has_segment_data:
            continue

        debug_log(f"[Segment Analysis] Appending '{extra_label}' row to notes sheet: {sheet_name}")

        row_data = [extra_label, extra_el]
        for col_key in sorted_role_cols:
            v = extra_vals.get(col_key)
            if v is not None:
                vc = unicodedata.normalize('NFKC', str(v)).replace(',', '').strip()
                try:
                    row_data.append(float(vc))
                except ValueError:
                    row_data.append(v)
            else:
                row_data.append("")

        ws.append(row_data)
        appended = True

    if not appended:
        return

    # 追加した行に書式を適用
    last_row = ws.max_row
    start_row = last_row - sum(
        1 for el, _ in _EXTRA_SEGMENT_ELEMENTS
        if global_element_period_values.get(el) and any(
            global_element_period_values[el].get(c) is not None
            for c in sorted_role_cols
            if len(c) == 3 and c[1] not in non_segment_dims
        )
    ) + 1
    for r in range(start_row, last_row + 1):
        for cell in ws[r][2:]:
            if isinstance(cell.value, (int, float)):
                cell.number_format = r'#,##0_ ;[Red]\-#,##0 '


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
    #    利益  : 「利益」または「損失」を含むラベルを全候補として収集
    # -----------------------------------------------------------------------
    target_sales_label = None
    profit_label_candidates = []
    for label in unique_labels_ordered:
        if target_sales_label is None and label == "計":
            target_sales_label = label
        if "利益" in label or "損失" in label:
            profit_label_candidates.append(label)

    debug_log(f"[PPM Analysis] max_year={max_year}, Sales label='{target_sales_label}', Profit candidates={profit_label_candidates}")

    # -----------------------------------------------------------------------
    # 3. 11年分の年度リスト（昇順: max_year-10 ～ max_year）
    # -----------------------------------------------------------------------
    NUM_YEARS = 11
    target_years = list(range(max_year - 10, max_year + 1))

    # 各年の analysis_ws 行番号（データなし年は None）
    # 利益は年度ごとに候補を順に試し、最初にデータがある行を採用する
    # （ラベル名が年度途中で変わった場合や行は存在するがデータ空の場合に対応）
    def _row_has_data(row):
        """analysis_ws の指定行（列3以降）に数値データが1つ以上あるか確認する。"""
        if row is None:
            return False
        for c in range(3, max_col + 1):
            v = analysis_ws.cell(row, c).value
            if isinstance(v, (int, float)):
                return True
        return False

    sales_src_rows = [lookup.get((target_sales_label, y)) if target_sales_label else None for y in target_years]
    profit_src_rows = []
    for y in target_years:
        src_row = None
        for candidate in profit_label_candidates:
            row = lookup.get((candidate, y))
            if row is not None and _row_has_data(row):
                src_row = row
                break
        profit_src_rows.append(src_row)

    # -----------------------------------------------------------------------
    # 3b. 列位置の検出（報告セグメント / 以外 / 及びその他の合計）
    # -----------------------------------------------------------------------
    def _read_numeric(ws, row, col):
        """セルの実数値を返す。SUM式セルは個別列を合計して代替する。"""
        if row is None:
            return None
        val = ws.cell(row, col).value
        if isinstance(val, (int, float)):
            return val
        if isinstance(val, str) and val.startswith('=SUM('):
            total = 0.0
            has_val = False
            for c in range(3, col):
                v = ws.cell(row, c).value
                if isinstance(v, (int, float)):
                    total += v
                    has_val = True
            return total if has_val else None
        return None

    hokoku_col = None   # 「報告セグメント」列（グラフ末端）
    igai_col   = None   # 「報告セグメント以外の全てのセグメント」列
    goukei_col = None   # 「報告セグメント及びその他の合計」列
    for _ci in range(3, max_col + 1):
        _hv = analysis_ws.cell(1, _ci).value
        if not _hv:
            continue
        _hv_str = str(_hv)
        if "報告セグメント" not in _hv_str:
            continue
        if "以外" in _hv_str:
            igai_col = _ci
        elif "及びその他" in _hv_str:
            goukei_col = _ci
        elif hokoku_col is None:
            hokoku_col = _ci
    if hokoku_col is None:
        # 「報告セグメント」単独列がない場合: 「以外」列の手前を末端とする
        hokoku_col = (igai_col - 1) if igai_col else max_col

    # 「報告セグメント及びその他の合計」列がなければ analysis_ws に追加
    if igai_col and goukei_col is None:
        new_col = max_col + 1
        analysis_ws.cell(1, new_col).value = "報告セグメント及びその他の合計"
        _h_letter = get_column_letter(hokoku_col)
        _i_letter = get_column_letter(igai_col)
        for _ri in range(2, analysis_ws.max_row + 1):
            if any(isinstance(analysis_ws.cell(_ri, c).value, (int, float))
                   for c in range(3, max_col + 1)):
                analysis_ws.cell(_ri, new_col).value = (
                    f"=SUM({_h_letter}{_ri},{_i_letter}{_ri})"
                )
        goukei_col = new_col
        max_col = new_col
        debug_log(f"[PPM Analysis] Added '報告セグメント及びその他の合計' column at col {new_col}")

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
                cur_val  = _read_numeric(analysis_ws, cur_src,  col_idx)
                prev_val = _read_numeric(analysis_ws, prev_src, col_idx)
                if cur_val is not None and prev_val is not None and prev_val != 0:
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
                s_val = _read_numeric(analysis_ws, s_src, col_idx)
                p_val = _read_numeric(analysis_ws, p_src, col_idx)
                if s_val is not None and p_val is not None and s_val != 0:
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
                year_sales.append(_read_numeric(analysis_ws, s_src, col_idx))
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
    LATEST_IDX     = NUM_YEARS - 1   # 最新年（インデックス10）
    FIVE_AGO_IDX   = NUM_YEARS - 6   # 5年前（インデックス5）
    FIVE_AGO_OFFSET = 5

    # chart_end_col: 「報告セグメント及びその他の合計」列（なければ「報告セグメント」列）
    # グラフには hokoku_col までしか使わないが、集約表はこの列まで表示する
    chart_end_col = goukei_col or hokoku_col

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

    def _valid_cols(year_idx):
        """3指標すべてが揃っている列のリストを返す（報告セグメント及びその他の合計列まで含む）"""
        result = []
        for ci in range(3, chart_end_col + 1):
            if (ci < len(profit_margins[year_idx])  and profit_margins[year_idx][ci]  is not None and
                ci < len(growth_rates[year_idx])    and growth_rates[year_idx][ci]    is not None and
                ci < len(sales_values[year_idx])    and sales_values[year_idx][ci]    is not None):
                result.append(ci)
        return result

    def _append_data_section(year_idx, sales_row, _profit_row, growth_row, margin_row):
        """集約データ4行（ヘッダ・利益率・成長率・売上）を追加して開始行を返す"""
        sec_start = ppm_ws.max_row + 1
        vcols = _valid_cols(year_idx)
        # グラフに使う列は hokoku_col まで（以外・合計列はグラフに含めない）
        vcols_chart = [ci for ci in vcols if ci <= hokoku_col]

        # ヘッダ行
        hrow = [f"=B{sales_row}", "セグメント名"]
        for ci in vcols:
            hrow.append(f"={get_column_letter(ci)}1")
        if vcols_chart and vcols_chart[-1] == hokoku_col:
            # 報告セグメント列のラベルを「計」に置き換える
            hrow[2 + vcols.index(hokoku_col)] = "計"
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

        # 売上行（hokoku_col のみバブルサイズ調整用に *1%）
        srow = [f"=B{sales_row}", f"=TRIM(A{sales_row})"]
        for ci in vcols:
            cl = get_column_letter(ci)
            srow.append(f"={cl}{sales_row}*1%" if ci == hokoku_col else f"={cl}{sales_row}")
        ppm_ws.append(srow)

        sec_end = ppm_ws.max_row
        sec_max_col = max(3, 2 + len(vcols))
        chart_sec_max_col = max(3, 2 + len(vcols_chart))

        # 書式
        for ci in range(3, sec_max_col + 1):
            cl = get_column_letter(ci)
            ppm_ws[f'{cl}{sec_start + 1}'].number_format = '0%'
            ppm_ws[f'{cl}{sec_start + 2}'].number_format = '0%'
            ppm_ws[f'{cl}{sec_start + 3}'].number_format = r'#,##0_);[Red](#,##0)'

        return sec_start, sec_end, sec_max_col, vcols, chart_sec_max_col

    # 最新年データ
    latest_sales_row  = sales_end_row
    latest_profit_row = profit_end_row
    latest_growth_row = growth_end_row
    latest_margin_row = margin_end_row

    # -----------------------------------------------------------------------
    # 9. 軸範囲計算（最新年 + 5年前 の共通スケール）
    # -----------------------------------------------------------------------
    def _axis_values(arr, year_indices):
        vals = []
        for yi in year_indices:
            if 0 <= yi < len(arr):
                for ci, v in enumerate(arr[yi]):
                    # 軸計算は hokoku_col まで（以外・合計列は除外）
                    if ci >= 3 and ci <= hokoku_col and v is not None:
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

    lat_start, lat_end, lat_max_col, _, lat_chart_max_col = _append_data_section(
        LATEST_IDX, latest_sales_row, latest_profit_row, latest_growth_row, latest_margin_row
    )

    chart_latest = _make_chart(_fmt_year_str(year_val_for_title))
    _add_series(chart_latest, ppm_ws, lat_start, lat_chart_max_col)

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

        five_start, five_end, five_max_col, _, five_chart_max_col = _append_data_section(
            FIVE_AGO_IDX, five_sales_row, five_profit_row, five_growth_row, five_margin_row
        )

        chart_5y = _make_chart(_fmt_year_str(y_val_5y))
        _add_series(chart_5y, ppm_ws, five_start, five_chart_max_col)
        debug_log("[PPM Analysis] Added 5-year comparison section")

    # -----------------------------------------------------------------------
    # 12. チャートをシートに配置
    # -----------------------------------------------------------------------
    chart_row = ppm_ws.max_row + 2
    ppm_ws.add_chart(chart_latest, f'B{chart_row}')
    if chart_5y:
        ppm_ws.add_chart(chart_5y, f'I{chart_row}')

    debug_log(f"[PPM Analysis] Completed PPM analysis sheet: {ppm_sheet_name}")


def _create_ppm_analysis_sheet_ifrs(workbook, analysis_sheet_name, used_sheet_names, debug_log):
    """
    IFRS用PPM分析シートを作成（内部関数）

    IFRSセグメント分析シートから売上収益・セグメント利益を検出し、
    日本基準PPM分析と同一フォーマットのシートを生成する。
    ラベル検出ロジックは _create_ebitda_sheet_ifrs と共通。
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

    debug_log(f"[PPM IFRS] Creating PPM analysis sheet: {ppm_sheet_name}")

    if analysis_sheet_name not in workbook.sheetnames:
        debug_log(f"[PPM IFRS] Analysis sheet '{analysis_sheet_name}' not found, skipping PPM sheet")
        return

    analysis_ws = workbook[analysis_sheet_name]
    ppm_ws = workbook.create_sheet(title=ppm_sheet_name)
    used_sheet_names.add(ppm_sheet_name)

    escaped_sheet_name = analysis_sheet_name.replace("'", "''")
    max_col = analysis_ws.max_column

    # -----------------------------------------------------------------------
    # 1. analysis_ws を走査して (ラベル, 年度) -> 行番号 のルックアップを構築
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
        debug_log("[PPM IFRS] No years found in analysis sheet, skipping PPM sheet")
        return

    # -----------------------------------------------------------------------
    # 2. IFRSラベル検出（_create_ebitda_sheet_ifrs と同一ロジック）
    #    売上収益: 「収益」or「売上」を含み、「外部顧客」「セグメント間」を除く
    #    利益: 「利益」を含むラベルを全候補として収集
    # -----------------------------------------------------------------------
    target_sales_label = None
    profit_label_candidates = []

    for label in unique_labels_ordered:
        if target_sales_label is None:
            if (("収益" in label or "売上" in label)
                    and "外部顧客" not in label
                    and "セグメント間" not in label):
                target_sales_label = label
        if "利益" in label:
            profit_label_candidates.append(label)

    debug_log(f"[PPM IFRS] max_year={max_year}, Sales label='{target_sales_label}', Profit candidates={profit_label_candidates}")

    # -----------------------------------------------------------------------
    # 3. 11年分の年度リスト（昇順: max_year-10 ～ max_year）
    # -----------------------------------------------------------------------
    NUM_YEARS = 11
    target_years = list(range(max_year - 10, max_year + 1))

    def _row_has_data(row):
        """analysis_ws の指定行（列3以降）に数値データが1つ以上あるか確認する。"""
        if row is None:
            return False
        for c in range(3, max_col + 1):
            v = analysis_ws.cell(row, c).value
            if isinstance(v, (int, float)):
                return True
        return False

    sales_src_rows = [lookup.get((target_sales_label, y)) if target_sales_label else None for y in target_years]
    # 利益は年度ごとに候補を順に試し、最初に実データがある行を採用する
    # （ラベル名が年度途中で変わった場合や行は存在するがデータ空の場合に対応）
    profit_src_rows = []
    for y in target_years:
        src_row = None
        for candidate in profit_label_candidates:
            row = lookup.get((candidate, y))
            if row is not None and _row_has_data(row):
                src_row = row
                break
        profit_src_rows.append(src_row)

    # -----------------------------------------------------------------------
    # 3b. 列位置の検出（報告セグメント / 以外 / 及びその他の合計）
    # -----------------------------------------------------------------------
    def _read_numeric(ws, row, col):
        """セルの実数値を返す。SUM式セルは個別列を合計して代替する。"""
        if row is None:
            return None
        val = ws.cell(row, col).value
        if isinstance(val, (int, float)):
            return val
        # SUM式セル（報告セグメント合計列など）は個別列を合計
        if isinstance(val, str) and val.startswith('=SUM('):
            total = 0.0
            has_val = False
            for c in range(3, col):
                v = ws.cell(row, c).value
                if isinstance(v, (int, float)):
                    total += v
                    has_val = True
            return total if has_val else None
        return None

    hokoku_col = None   # 「報告セグメント（合計）」列（グラフ末端）
    igai_col   = None   # 「報告セグメント以外の全てのセグメント」列
    goukei_col = None   # 「報告セグメント及びその他の合計」列
    for _ci in range(3, max_col + 1):
        _hv = analysis_ws.cell(1, _ci).value
        if not _hv:
            continue
        _hv_str = str(_hv)
        if "報告セグメント" not in _hv_str:
            continue
        if "以外" in _hv_str:
            igai_col = _ci
        elif "及びその他" in _hv_str:
            goukei_col = _ci
        elif hokoku_col is None:
            hokoku_col = _ci
    if hokoku_col is None:
        # 「報告セグメント」単独列がない場合: 「以外」列の手前を末端とする
        hokoku_col = (igai_col - 1) if igai_col else max_col

    # 「報告セグメント及びその他の合計」列がなければ analysis_ws に追加
    if igai_col and goukei_col is None:
        new_col = max_col + 1
        analysis_ws.cell(1, new_col).value = "報告セグメント及びその他の合計"
        _h_letter = get_column_letter(hokoku_col)
        _i_letter = get_column_letter(igai_col)
        for _ri in range(2, analysis_ws.max_row + 1):
            if any(isinstance(analysis_ws.cell(_ri, c).value, (int, float))
                   for c in range(3, max_col + 1)):
                analysis_ws.cell(_ri, new_col).value = (
                    f"=SUM({_h_letter}{_ri},{_i_letter}{_ri})"
                )
        goukei_col = new_col
        max_col = new_col
        debug_log(f"[PPM IFRS] Added '報告セグメント及びその他の合計' column at col {new_col}")

    # chart_end_col: 「報告セグメント及びその他の合計」列（なければ「報告セグメント」列）
    chart_end_col = goukei_col or hokoku_col

    # -----------------------------------------------------------------------
    # 4. ppm_ws の構築（日本基準PPMと同一）
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
        data_row = ["　売上収益"]
        for col_idx in range(2, max_col + 1):
            cl = get_column_letter(col_idx)
            if col_idx == 2:
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
            if col_idx == 2:
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
    growth_rates = []

    for idx in range(NUM_YEARS):
        year_row_ref = sales_start_row + idx
        growth_row = ["売上収益対前年増加率"]
        year_growth_rates = [None, None]

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
                cur_src  = sales_src_rows[idx]
                prev_src = sales_src_rows[idx - 1]
                cur_val  = _read_numeric(analysis_ws, cur_src,  col_idx)
                prev_val = _read_numeric(analysis_ws, prev_src, col_idx)
                if cur_val is not None and prev_val is not None and prev_val != 0:
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
    profit_margins = []

    for idx in range(NUM_YEARS):
        year_row_ref = sales_start_row + idx
        margin_row = ["売上高利益率"]
        year_profit_margins = [None, None]

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
                s_src = sales_src_rows[idx]
                p_src = profit_src_rows[idx]
                s_val = _read_numeric(analysis_ws, s_src, col_idx)
                p_val = _read_numeric(analysis_ws, p_src, col_idx)
                if s_val is not None and p_val is not None and s_val != 0:
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
                year_sales.append(_read_numeric(analysis_ws, s_src, col_idx))
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

    LATEST_IDX     = NUM_YEARS - 1
    FIVE_AGO_IDX   = NUM_YEARS - 6
    FIVE_AGO_OFFSET = 5

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

    def _valid_cols(year_idx):
        """3指標すべてが揃っている列のリストを返す（報告セグメント及びその他の合計列まで含む）"""
        result = []
        for ci in range(3, chart_end_col + 1):
            if (ci < len(profit_margins[year_idx])  and profit_margins[year_idx][ci]  is not None and
                ci < len(growth_rates[year_idx])    and growth_rates[year_idx][ci]    is not None and
                ci < len(sales_values[year_idx])    and sales_values[year_idx][ci]    is not None):
                result.append(ci)
        return result

    def _append_data_section(year_idx, sales_row, growth_row, margin_row):
        sec_start = ppm_ws.max_row + 1
        vcols = _valid_cols(year_idx)
        # グラフに使う列は hokoku_col まで（以外・合計列はグラフに含めない）
        vcols_chart = [ci for ci in vcols if ci <= hokoku_col]

        hrow = [f"=B{sales_row}", "セグメント名"]
        for ci in vcols:
            hrow.append(f"={get_column_letter(ci)}1")
        if vcols_chart and vcols_chart[-1] == hokoku_col:
            hrow[2 + vcols.index(hokoku_col)] = "計"
        ppm_ws.append(hrow)

        mrow = [f"=B{margin_row}", f"=A{margin_row}"]
        for ci in vcols:
            mrow.append(f"={get_column_letter(ci)}{margin_row}")
        ppm_ws.append(mrow)

        grow = [f"=B{growth_row}", f"=A{growth_row}"]
        for ci in vcols:
            grow.append(f"={get_column_letter(ci)}{growth_row}")
        ppm_ws.append(grow)

        srow = [f"=B{sales_row}", f"=TRIM(A{sales_row})"]
        for ci in vcols:
            cl = get_column_letter(ci)
            srow.append(f"={cl}{sales_row}*1%" if ci == hokoku_col else f"={cl}{sales_row}")
        ppm_ws.append(srow)

        sec_end = ppm_ws.max_row
        sec_max_col = max(3, 2 + len(vcols))
        chart_sec_max_col = max(3, 2 + len(vcols_chart))

        for ci in range(3, sec_max_col + 1):
            cl = get_column_letter(ci)
            ppm_ws[f'{cl}{sec_start + 1}'].number_format = '0%'
            ppm_ws[f'{cl}{sec_start + 2}'].number_format = '0%'
            ppm_ws[f'{cl}{sec_start + 3}'].number_format = r'#,##0_);[Red](#,##0)'

        return sec_start, sec_end, sec_max_col, vcols, chart_sec_max_col

    latest_sales_row  = sales_end_row
    latest_growth_row = growth_end_row
    latest_margin_row = margin_end_row

    lat_start, lat_end, lat_max_col, _, lat_chart_max_col = _append_data_section(
        LATEST_IDX, latest_sales_row, latest_growth_row, latest_margin_row
    )

    # -----------------------------------------------------------------------
    # 9. 軸範囲計算
    # -----------------------------------------------------------------------
    def _axis_values(arr, year_indices):
        vals = []
        for yi in year_indices:
            if 0 <= yi < len(arr):
                for ci, v in enumerate(arr[yi]):
                    # 軸計算は hokoku_col まで（以外・合計列は除外）
                    if ci >= 3 and ci <= hokoku_col and v is not None:
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
        chart.y_axis.title = "売上収益対前年増加率"
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
    _add_series(chart_latest, ppm_ws, lat_start, lat_chart_max_col)

    # -----------------------------------------------------------------------
    # 11. 5年前データセクション
    # -----------------------------------------------------------------------
    ppm_ws.append([""] * max_col)

    chart_5y = None
    if NUM_YEARS > FIVE_AGO_OFFSET:
        five_sales_row  = sales_end_row  - FIVE_AGO_OFFSET
        five_growth_row = growth_end_row - FIVE_AGO_OFFSET
        five_margin_row = margin_end_row - FIVE_AGO_OFFSET

        y_val_5y = (analysis_ws.cell(sales_src_rows[FIVE_AGO_IDX], 2).value
                    if sales_src_rows[FIVE_AGO_IDX] else target_years[FIVE_AGO_IDX])

        five_start, five_end, five_max_col, _, five_chart_max_col = _append_data_section(
            FIVE_AGO_IDX, five_sales_row, five_growth_row, five_margin_row
        )

        chart_5y = _make_chart(_fmt_year_str(y_val_5y))
        _add_series(chart_5y, ppm_ws, five_start, five_chart_max_col)
        debug_log("[PPM IFRS] Added 5-year comparison section")

    # -----------------------------------------------------------------------
    # 12. チャートをシートに配置
    # -----------------------------------------------------------------------
    chart_row = ppm_ws.max_row + 2
    ppm_ws.add_chart(chart_latest, f'B{chart_row}')
    if chart_5y:
        ppm_ws.add_chart(chart_5y, f'I{chart_row}')

    debug_log(f"[PPM IFRS] Completed PPM analysis sheet: {ppm_sheet_name}")


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


def _create_ebitda_sheet_ifrs(workbook, analysis_sheet_name, used_sheet_names, debug_log):
    """
    IFRS用EBITDA分析シートを作成（内部関数）

    IFRSセグメント分析シートから売上収益・セグメント利益・償却費等を検出し、
    EBITDAを計算するシートを生成する。

    Args:
        workbook: openpyxlワークブック
        analysis_sheet_name: 参照元の分析シート名（例: 連結_セグメント情報等(IFRS)_分析）
        used_sheet_names: 使用済みシート名のセット
        debug_log: デバッグログ関数
    """
    import re
    from openpyxl.utils import get_column_letter

    ebitda_sheet_name = analysis_sheet_name + "_EBITDA"
    if len(ebitda_sheet_name) > 31:
        ebitda_sheet_name = analysis_sheet_name[:24] + "_EBITDA"

    debug_log(f"[EBITDA IFRS] Creating EBITDA sheet: {ebitda_sheet_name}")

    if analysis_sheet_name not in workbook.sheetnames:
        debug_log(f"[EBITDA IFRS] Analysis sheet '{analysis_sheet_name}' not found, skipping")
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
        debug_log("[EBITDA IFRS] No years found in analysis sheet, skipping")
        return

    # -----------------------------------------------------------------------
    # 2. 対象ラベルを検索
    # -----------------------------------------------------------------------
    target_sales_label = None    # 売上収益・営業収益・収益・売上高
    profit_label_candidates = [] # 全「利益」ヒット（複数年度でラベル変更に対応）
    refs_labels = []             # 償却費・償却額・減損損失（戻入除く）
    refs_minus_labels = []       # 減損の戻入・負ののれん

    for label in unique_labels_ordered:
        # 売上高: 「収益」or「売上」を含み、「外部顧客」「セグメント間」を除く
        if target_sales_label is None:
            if (("収益" in label or "売上" in label)
                    and "外部顧客" not in label
                    and "セグメント間" not in label):
                target_sales_label = label

        # セグメント利益: 「利益」を含む全候補を収集
        if "利益" in label:
            profit_label_candidates.append(label)

        # 減損の戻入（「戻入」or「戻し入」を含む）・負ののれん → refs_minus
        if (("減損" in label and ("戻入" in label or "戻し入" in label))
                or "負ののれん" in label):
            if label not in refs_minus_labels:
                refs_minus_labels.append(label)
        # 償却費・償却額・減損損失（戻入なし）→ refs
        elif (("償却費" in label or "償却額" in label or "減損損失" in label)
              and "戻入" not in label and "戻し入" not in label):
            if label not in refs_labels:
                refs_labels.append(label)

    debug_log(f"[EBITDA IFRS] max_year={max_year}, Sales='{target_sales_label}', "
              f"Profit candidates={profit_label_candidates}, refs={refs_labels}, refs_minus={refs_minus_labels}")

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

    def _first_valid_row(label):
        """ラベル表示用に最新年から最初の有効な行番号を返す"""
        for y in reversed(target_years):
            r = lookup.get((label, y))
            if r is not None:
                return r
        return None

    def _row_has_data(row):
        """analysis_ws の指定行（列3以降）に数値データが1つ以上あるか確認する。"""
        if row is None:
            return False
        for c in range(3, max_col + 1):
            v = analysis_ws.cell(row, c).value
            if isinstance(v, (int, float)):
                return True
        return False

    def _write_data_block_ifrs(label, src_rows_list, year_fallbacks=None):
        """セル参照ラベルでデータブロックを書き込み、ブロック開始行を返す。
        year_fallbacks: src_rowがNoneのとき年度列に表示するリテラル値のリスト（Noneなら空欄）"""
        block_start = ebitda_ws.max_row + 1
        label_row = _first_valid_row(label)
        for idx, src_row in enumerate(src_rows_list):
            label_formula = f"='{escaped}'!A{label_row}" if label_row is not None else label
            row_data = [label_formula]
            for col_idx in range(2, max_col + 1):
                cl = get_column_letter(col_idx)
                if col_idx == 2:
                    if src_row:
                        formula = f"='{escaped}'!B{src_row}"
                    elif year_fallbacks is not None:
                        formula = year_fallbacks[idx]
                    else:
                        formula = ""
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
    sales_block_start = _write_data_block_ifrs(
        target_sales_label or "売上高", sales_src_rows, year_fallbacks=target_years
    )

    # -----------------------------------------------------------------------
    # セグメント利益ブロック（複数ラベル対応）
    # 年度ごとに最初にデータがある候補が担当する。
    # 各候補を別ブロックとして出力し、担当外の年はデータ空（年度は表示）。
    # -----------------------------------------------------------------------
    # 年度ごとの担当ラベルと行番号を決定
    profit_assignment = {}  # year -> (label, row)
    for y in target_years:
        for candidate in profit_label_candidates:
            row = lookup.get((candidate, y))
            if row is not None and _row_has_data(row):
                profit_assignment[y] = (candidate, row)
                break

    # 利益候補ごとにブロックを作成（担当年のみ実データ行を渡す）
    profit_block_starts = []
    for candidate in profit_label_candidates:
        candidate_src_rows = []
        has_any = False
        for y in target_years:
            assignment = profit_assignment.get(y)
            if assignment is not None and assignment[0] == candidate:
                candidate_src_rows.append(assignment[1])
                has_any = True
            else:
                candidate_src_rows.append(None)
        if has_any:
            profit_block_starts.append(
                _write_data_block_ifrs(candidate, candidate_src_rows, year_fallbacks=target_years)
            )

    if not profit_block_starts:
        # 利益ラベルが全くない場合はダミー行を1ブロック追加
        dummy_rows = [None] * NUM_YEARS
        profit_block_starts.append(
            _write_data_block_ifrs("セグメント利益", dummy_rows, year_fallbacks=target_years)
        )

    # refsブロック
    refs_block_starts = []
    for lbl in refs_labels:
        src_rows = [lookup.get((lbl, y)) for y in target_years]
        refs_block_starts.append(_write_data_block_ifrs(lbl, src_rows, year_fallbacks=target_years))

    # refs_minusブロック
    refs_minus_block_starts = []
    for lbl in refs_minus_labels:
        src_rows = [lookup.get((lbl, y)) for y in target_years]
        refs_minus_block_starts.append(_write_data_block_ifrs(lbl, src_rows, year_fallbacks=target_years))

    # --- 空行 ---
    ebitda_ws.append([""] * max_col)

    # --- EBITDAブロック ---
    # 複数利益ブロックは年度ごとに排他的（担当年のみデータ）なので合算してよい。
    # profit_sum = IF(p1="",0,p1)+IF(p2="",0,p2)+...
    # any_profit_cond = OR(p1<>"", p2<>"", ...)
    ebitda_block_start = ebitda_ws.max_row + 1
    for idx in range(NUM_YEARS):
        row_data = ["EBITDA"]
        row_data.append(f"=B{sales_block_start + idx}")
        for col_idx in range(3, max_col + 1):
            cl = get_column_letter(col_idx)
            profit_cells = [f"{cl}{bs + idx}" for bs in profit_block_starts]
            refs_cells = [f"{cl}{bs + idx}" for bs in refs_block_starts]
            refs_minus_cells = [f"{cl}{bs + idx}" for bs in refs_minus_block_starts]

            # 利益の合計式（空白を0として加算）
            if len(profit_cells) == 1:
                profit_sum = profit_cells[0]
                any_profit_cond = f"{profit_cells[0]}<>\"\""
            else:
                profit_sum = "+".join(f"IF({p}=\"\",0,{p})" for p in profit_cells)
                any_profit_cond = "OR(" + ",".join(f"{p}<>\"\"" for p in profit_cells) + ")"

            if not refs_cells and not refs_minus_cells:
                formula = f"=IF(NOT({any_profit_cond}),\"\",{profit_sum})"
            elif not refs_minus_cells:
                refs_str = ",".join(refs_cells)
                formula = (
                    f"=IF(OR(NOT({any_profit_cond}),COUNT({refs_str})=0),"
                    f"\"\",{profit_sum}+SUM({refs_str}))"
                )
            elif not refs_cells:
                refs_minus_str = ",".join(refs_minus_cells)
                formula = (
                    f"=IF(OR(NOT({any_profit_cond}),COUNT({refs_minus_str})=0),"
                    f"\"\",{profit_sum}-SUM({refs_minus_str}))"
                )
            else:
                refs_str = ",".join(refs_cells)
                refs_minus_str = ",".join(refs_minus_cells)
                formula = (
                    f"=IF(OR(NOT({any_profit_cond}),COUNT({refs_str})=0),"
                    f"\"\",{profit_sum}+SUM({refs_str})-SUM({refs_minus_str}))"
                )
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
    for row in ebitda_ws.iter_rows(min_row=2, max_row=ebitda_block_end, min_col=3, max_col=max_col):
        for cell in row:
            cell.number_format = r'#,##0_ ;[Red]\-#,##0 '

    for row in ebitda_ws.iter_rows(min_row=ratio_block_start, max_row=ebitda_ws.max_row, min_col=3, max_col=max_col):
        for cell in row:
            cell.number_format = '0%'

    ebitda_ws.freeze_panes = 'B2'
    ebitda_ws.column_dimensions['A'].width = 31
    for col_idx in range(2, max_col + 1):
        ebitda_ws.column_dimensions[get_column_letter(col_idx)].width = 12

    debug_log(f"[EBITDA IFRS] Completed EBITDA sheet: {ebitda_sheet_name}")
