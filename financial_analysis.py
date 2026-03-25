"""
Financial Analysis Module for XBRL to Excel Conversion

ROE分析やその他の財務分析機能を提供するモジュール
"""

import openpyxl.utils


def create_roe_analysis_sheet(workbook, source_sheet_name, debug_log=None):
    """
    主要な経営指標等の推移シートからROE分析シートを生成

    Args:
        workbook: openpyxlワークブック
        source_sheet_name: 元シート名（例: "主要な経営指標等の推移（連結）(日本基準)"）
        debug_log: デバッグログ関数（オプション）
    """
    # デバッグログ関数がない場合はダミー関数を使用
    if debug_log is None:
        def debug_log(msg):
            pass

    if source_sheet_name not in workbook.sheetnames:
        return

    source_ws = workbook[source_sheet_name]
    analysis_sheet_name = f"{source_sheet_name}_ROE分析"

    # 既存の分析シートがあれば削除
    if analysis_sheet_name in workbook.sheetnames:
        workbook.remove(workbook[analysis_sheet_name])

    # 新しいシートを作成
    analysis_ws = workbook.create_sheet(analysis_sheet_name)

    # 列数を取得
    num_cols = source_ws.max_column

    # 参照する行番号を特定（元シートから）
    row_mapping = {}  # 英語名 -> 行番号のマッピング
    for row in range(2, source_ws.max_row + 1):
        english_name = source_ws.cell(row, 2).value  # B列: 項目（英名）
        if english_name:
            row_mapping[english_name] = row

    # 必要な勘定科目の英語名
    sales_key = 'jpcrp_cor_NetSalesSummaryOfBusinessResults'  # 売上高
    profit_key = 'jpcrp_cor_ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults'  # 親会社株主に帰属する当期純利益
    net_assets_key = 'jpcrp_cor_NetAssetsSummaryOfBusinessResults'  # 純資産額
    total_assets_key = 'jpcrp_cor_TotalAssetsSummaryOfBusinessResults'  # 総資産額
    equity_ratio_key = 'jpcrp_cor_EquityToAssetRatioSummaryOfBusinessResults'  # 自己資本比率
    roe_key = 'jpcrp_cor_RateOfReturnOnEquitySummaryOfBusinessResults'  # 自己資本利益率

    # 元シートの行番号を取得
    sales_row = row_mapping.get(sales_key)
    profit_row = row_mapping.get(profit_key)
    net_assets_row = row_mapping.get(net_assets_key)
    total_assets_row = row_mapping.get(total_assets_key)
    equity_ratio_row = row_mapping.get(equity_ratio_key)
    roe_row = row_mapping.get(roe_key)

    # 売上高が見つからない場合は、セクションヘッダーの次の行を使用
    if sales_row is None:
        header_key = 'jpcrp_cor_BusinessResultsOfGroupHeading'
        header_row_num = row_mapping.get(header_key)
        if header_row_num:
            # ヘッダーの次の行を売上高として使用
            sales_row = header_row_num + 1
            debug_log(f"Sales row not found, using first item after header (row {sales_row})")

    # ROE分析に必要な項目がすべて存在するかチェック
    required_items = {
        '売上高': sales_row,
        '当期純利益': profit_row,
        '総資産額': total_assets_row,
        '自己資本比率': equity_ratio_row,
        'ROE': roe_row
    }
    missing_items = [name for name, row in required_items.items() if row is None]
    if missing_items:
        debug_log(f"ROE analysis skipped for '{source_sheet_name}': missing items: {', '.join(missing_items)}")
        return

    # ヘッダー行と基本指標をすべて元シートから参照する形で追加
    def add_reference_row_full(source_row_num):
        """元シートの行全体を参照する行を追加"""
        row_data = []
        for col in range(1, num_cols + 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            # 元シートのセルを参照する数式
            formula = f"='{source_sheet_name}'!{col_letter}{source_row_num}"
            row_data.append(formula)
        analysis_ws.append(row_data)

    # 1行目: ヘッダー（勘定科目、項目（英名）など）
    add_reference_row_full(1)

    # 2行目: セクションヘッダー（連結経営指標等）
    header_key = 'jpcrp_cor_BusinessResultsOfGroupHeading'
    header_row_num = row_mapping.get(header_key)
    if header_row_num:
        add_reference_row_full(header_row_num)

    # 3-8行目: 基本指標（元シートから参照）
    add_reference_row_full(sales_row)
    add_reference_row_full(profit_row)
    add_reference_row_full(net_assets_row)
    add_reference_row_full(total_assets_row)
    add_reference_row_full(equity_ratio_row)
    add_reference_row_full(roe_row)

    # 現在の行番号を記録（上記で追加した行の位置）
    current_row = analysis_ws.max_row
    sales_analysis_row = current_row - 5
    profit_analysis_row = current_row - 4
    net_assets_analysis_row = current_row - 3
    total_assets_analysis_row = current_row - 2
    equity_ratio_analysis_row = current_row - 1
    roe_analysis_row = current_row

    # 空行
    analysis_ws.append([''] * num_cols)

    # 計算指標
    # 自己資本 = 総資産額 × 自己資本比率
    equity_row_num = analysis_ws.max_row + 1
    equity_row = ['　　　自己資本', '']
    for col in range(3, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        formula = f"={col_letter}{total_assets_analysis_row}*{col_letter}{equity_ratio_analysis_row}"
        equity_row.append(formula)
    analysis_ws.append(equity_row)

    # 自己資本（平均） = AVERAGE(前期:当期)
    equity_avg_row_num = analysis_ws.max_row + 1
    equity_avg_row = ['　　　自己資本（平均）', '']
    # 最初の期（C列）は計算不可
    equity_avg_row.append('')
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        prev_col_letter = openpyxl.utils.get_column_letter(col - 1)
        formula = f"=AVERAGE({prev_col_letter}{equity_row_num}:{col_letter}{equity_row_num})"
        equity_avg_row.append(formula)
    analysis_ws.append(equity_avg_row)

    # 総資産（平均） = AVERAGE(前期:当期)
    total_assets_avg_row_num = analysis_ws.max_row + 1
    total_assets_avg_row = ['　　　総資産（平均）', '']
    # 最初の期（C列）は計算不可
    total_assets_avg_row.append('')
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        prev_col_letter = openpyxl.utils.get_column_letter(col - 1)
        formula = f"=AVERAGE({prev_col_letter}{total_assets_analysis_row}:{col_letter}{total_assets_analysis_row})"
        total_assets_avg_row.append(formula)
    analysis_ws.append(total_assets_avg_row)

    # 空行
    analysis_ws.append([''] * num_cols)

    # ROE分析指標
    # ROE = 元シートのROE
    roe_calc_row_num = analysis_ws.max_row + 1
    roe_calc_row = ['　　　自己資本利益率(ROE)', '']
    # 最初の期（C列）は計算不可（平均が必要なため）
    roe_calc_row.append('')
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        formula = f"={col_letter}{roe_analysis_row}"
        roe_calc_row.append(formula)
    analysis_ws.append(roe_calc_row)

    # ROS = 当期純利益 / 売上高
    ros_row_num = analysis_ws.max_row + 1
    ros_row = ['　　　売上高利益率(ROS)', '']
    ros_row.append('')  # 最初の期
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        formula = f"={col_letter}{profit_analysis_row}/{col_letter}{sales_analysis_row}"
        ros_row.append(formula)
    analysis_ws.append(ros_row)

    # TOR = 売上高 / 総資産（平均）
    tor_row_num = analysis_ws.max_row + 1
    tor_row = ['　　　総資産回転率(TOR)', '']
    tor_row.append('')  # 最初の期
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        formula = f"={col_letter}{sales_analysis_row}/{col_letter}{total_assets_avg_row_num}"
        tor_row.append(formula)
    analysis_ws.append(tor_row)

    # LRV = 総資産（平均） / 自己資本（平均）
    lrv_row_num = analysis_ws.max_row + 1
    lrv_row = ['　　　レバレッジ(LRV)', '']
    lrv_row.append('')  # 最初の期
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        formula = f"={col_letter}{total_assets_avg_row_num}/{col_letter}{equity_avg_row_num}"
        lrv_row.append(formula)
    analysis_ws.append(lrv_row)

    # 検算1: ROS * TOR * LRV = ROE
    check1_row_num = analysis_ws.max_row + 1
    check1_row = ['　　　検算1(ROS*TOR*LEV=ROE)', '']
    check1_row.append('')  # 最初の期
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        formula = f"=PRODUCT({col_letter}{ros_row_num}:{col_letter}{lrv_row_num})"
        check1_row.append(formula)
    analysis_ws.append(check1_row)

    # 検算2: 検算1 = ROE（TRUE/FALSE）
    check2_row_num = analysis_ws.max_row + 1
    check2_row = ['　　　検算2(検算1=ROE)', '']
    check2_row.append('')  # 最初の期
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        formula = f"=ROUND({col_letter}{roe_calc_row_num},1)=ROUND({col_letter}{check1_row_num},1)"
        check2_row.append(formula)
    analysis_ws.append(check2_row)

    # ROA = 当期純利益 / 総資産（平均）
    roa_row_num = analysis_ws.max_row + 1
    roa_row = ['　　　ROA(総資産利益率)', '']
    roa_row.append('')  # 最初の期
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        formula = f"={col_letter}{profit_analysis_row}/{col_letter}{total_assets_avg_row_num}"
        roa_row.append(formula)
    analysis_ws.append(roa_row)

    # ============================================================================
    # セルの表示形式を設定
    # ============================================================================
    # 数値フォーマット定義
    number_format_integer = r'#,##0_ ;[Red]\-#,##0\ '  # 整数（カンマ区切り）
    number_format_decimal = r'#,##0_);[Red](#,##0)'  # 整数（カンマ区切り、負数は括弧）
    number_format_decimal2 = r'#,##0.00_);[Red](#,##0.00)'  # 小数2桁
    number_format_percent = r'0.0%'  # パーセント（小数1桁）

    # 基本指標の表示形式（元シートから参照している行）
    for col in range(3, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)

        # 売上高、当期純利益、純資産額、総資産額: 整数カンマ区切り
        for row_num in [sales_analysis_row, profit_analysis_row, net_assets_analysis_row, total_assets_analysis_row]:
            analysis_ws[f'{col_letter}{row_num}'].number_format = number_format_integer

        # 自己資本比率、自己資本利益率: パーセント
        for row_num in [equity_ratio_analysis_row, roe_analysis_row]:
            analysis_ws[f'{col_letter}{row_num}'].number_format = number_format_percent

    # 計算指標の表示形式
    for col in range(3, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)

        # 自己資本、自己資本（平均）、総資産（平均）: 整数カンマ区切り（括弧）
        for row_num in [equity_row_num, equity_avg_row_num, total_assets_avg_row_num]:
            analysis_ws[f'{col_letter}{row_num}'].number_format = number_format_decimal

    # ROE分析指標の表示形式
    for col in range(4, num_cols + 1):  # D列から（C列は空）
        col_letter = openpyxl.utils.get_column_letter(col)

        # ROE、ROS、ROA: パーセント
        for row_num in [roe_calc_row_num, ros_row_num, roa_row_num]:
            analysis_ws[f'{col_letter}{row_num}'].number_format = number_format_percent

        # TOR、LRV: 小数2桁
        for row_num in [tor_row_num, lrv_row_num]:
            analysis_ws[f'{col_letter}{row_num}'].number_format = number_format_decimal2

        # 検算1: パーセント
        analysis_ws[f'{col_letter}{check1_row_num}'].number_format = number_format_percent

        # 検算2: パーセント（実際はTRUE/FALSEだが元シートに合わせる）
        analysis_ws[f'{col_letter}{check2_row_num}'].number_format = number_format_percent

    # ============================================================================
    # 列幅の設定とウィンドウ枠の固定
    # ============================================================================
    # A列: 幅28
    analysis_ws.column_dimensions['A'].width = 28

    # B列: 非表示
    analysis_ws.column_dimensions['B'].hidden = True

    # C列以降（年度の列）: 幅12
    for col in range(3, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        analysis_ws.column_dimensions[col_letter].width = 12

    # B2でウィンドウ枠を固定
    analysis_ws.freeze_panes = 'B2'

    # ============================================================================
    # 対前年増加率セクション（A23以降）
    # ============================================================================
    # 空行（2行）
    analysis_ws.append([''] * num_cols)
    analysis_ws.append([''] * num_cols)

    # A23行: "　対前年増加率" ヘッダー
    yoy_header_row_num = analysis_ws.max_row + 1
    yoy_header_row = ['　対前年増加率', '']
    # C列とD列は空欄
    yoy_header_row.append('')  # C列
    yoy_header_row.append('')  # D列
    # E列以降は1行目を参照する数式
    for col in range(5, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        yoy_header_row.append(f'={col_letter}1')
    analysis_ws.append(yoy_header_row)

    # 対前年増加率を計算する行の定義
    # 行24-29: 基本指標（売上高、当期純利益、純資産額、総資産額、自己資本比率、ROE）
    yoy_rows_basic = []
    for source_row in [sales_analysis_row, profit_analysis_row, net_assets_analysis_row,
                       total_assets_analysis_row, equity_ratio_analysis_row, roe_analysis_row]:
        yoy_row_num = analysis_ws.max_row + 1
        yoy_rows_basic.append(yoy_row_num)
        yoy_row = [f'=A{source_row}', '']  # A列は元の行を参照
        yoy_row.append('')  # C列は空（初年度は前年がない）
        # D列以降: =D{source_row}/C{source_row}-1
        for col in range(4, num_cols + 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            prev_col_letter = openpyxl.utils.get_column_letter(col - 1)
            formula = f"={col_letter}{source_row}/{prev_col_letter}{source_row}-1"
            yoy_row.append(formula)
        analysis_ws.append(yoy_row)

    # 空行
    analysis_ws.append([''] * num_cols)

    # 行31-33: 計算指標（自己資本、自己資本（平均）、総資産（平均））
    yoy_rows_calc = []
    for source_row in [equity_row_num, equity_avg_row_num, total_assets_avg_row_num]:
        yoy_row_num = analysis_ws.max_row + 1
        yoy_rows_calc.append(yoy_row_num)
        yoy_row = [f'=A{source_row}', '']
        yoy_row.append('')  # C列は空
        for col in range(4, num_cols + 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            prev_col_letter = openpyxl.utils.get_column_letter(col - 1)
            formula = f"={col_letter}{source_row}/{prev_col_letter}{source_row}-1"
            yoy_row.append(formula)
        analysis_ws.append(yoy_row)

    # 空行
    analysis_ws.append([''] * num_cols)

    # 行35-41: ROE分析指標
    yoy_rows_roe = []
    for source_row in [roe_calc_row_num, ros_row_num, tor_row_num, lrv_row_num,
                       check1_row_num, check2_row_num, roa_row_num]:
        yoy_row_num = analysis_ws.max_row + 1
        yoy_rows_roe.append(yoy_row_num)
        yoy_row = [f'=A{source_row}', '']
        yoy_row.append('')  # C列は空
        for col in range(4, num_cols + 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            prev_col_letter = openpyxl.utils.get_column_letter(col - 1)
            formula = f"={col_letter}{source_row}/{prev_col_letter}{source_row}-1"
            yoy_row.append(formula)
        analysis_ws.append(yoy_row)

    # 対前年増加率セクションの表示形式を設定（パーセント）
    for col in range(4, num_cols + 1):
        col_letter = openpyxl.utils.get_column_letter(col)
        for row_num in yoy_rows_basic + yoy_rows_calc + yoy_rows_roe:
            analysis_ws[f'{col_letter}{row_num}'].number_format = number_format_percent

    # ============================================================================
    # 10年前からの増加率計算（Q列）
    # ============================================================================
    # 最新の年の列を特定（最後のデータ列）
    latest_col = num_cols
    latest_col_letter = openpyxl.utils.get_column_letter(latest_col)

    # 10年前の列を特定（基準は3列目から）
    # 基準となる5つの指標がすべて揃っているかチェック
    target_rows = [sales_analysis_row, profit_analysis_row, net_assets_analysis_row,
                   total_assets_analysis_row, equity_ratio_analysis_row]

    # 10年前（11列前）のデータがあるかチェック（最古は3列目=C列）
    base_col = None
    ten_years_ago_col = latest_col - 10  # 10年前の列

    if ten_years_ago_col >= 3:  # C列以降であればチェック
        # 10年前のデータが5つの指標すべてで揃っているかチェック
        all_data_exists = True
        ten_years_col_letter = openpyxl.utils.get_column_letter(ten_years_ago_col)

        for row_num in target_rows:
            cell_formula = analysis_ws[f'{ten_years_col_letter}{row_num}'].value
            # 空欄や数式でない場合はデータなしと判断
            if not cell_formula:
                all_data_exists = False
                break

        if all_data_exists:
            base_col = ten_years_ago_col
            debug_log(f"Using 10-year comparison: column {ten_years_col_letter}")

    # 10年前のデータがない場合は、最長の年で計算
    if base_col is None:
        # 3列目（C列）から順にチェックして、5つの指標がすべて揃っている最古の列を探す
        for col in range(3, latest_col):
            col_letter = openpyxl.utils.get_column_letter(col)
            all_data_exists = True

            for row_num in target_rows:
                cell_formula = analysis_ws[f'{col_letter}{row_num}'].value
                if not cell_formula:
                    all_data_exists = False
                    break

            if all_data_exists:
                base_col = col
                debug_log(f"Using longest available period: column {col_letter}")
                break

    # 増加率列を追加（Q列 = num_cols + 1）
    if base_col is not None:
        growth_col = num_cols + 1
        growth_col_letter = openpyxl.utils.get_column_letter(growth_col)
        base_col_letter = openpyxl.utils.get_column_letter(base_col)

        # Q1: 期間表示（例: "2024/2014"）
        period_formula = f"=YEAR({latest_col_letter}1) & \"/\" & YEAR({base_col_letter}1)"
        analysis_ws[f'{growth_col_letter}1'] = period_formula

        # Q2: 空（ヘッダー行）
        analysis_ws[f'{growth_col_letter}2'] = ''

        # Q3-Q20: 空白（年平均増加率は対前年増加率セクションのQ24-Q41に移動）

        # Q23: "　対前年増加率" ヘッダー（Q1を参照）
        analysis_ws[f'{growth_col_letter}{yoy_header_row_num}'] = f'={growth_col_letter}1'

        # Q24-Q41: 対前年増加率セクションの年平均増加率（Q3-Q20を移動）
        # Q24-Q29: 基本指標
        for idx, row_num in enumerate(yoy_rows_basic):
            source_cagr_row = [sales_analysis_row, profit_analysis_row, net_assets_analysis_row,
                             total_assets_analysis_row, equity_ratio_analysis_row, roe_analysis_row][idx]
            cagr_formula = (f"=({latest_col_letter}{source_cagr_row}/{base_col_letter}{source_cagr_row})"
                          f"^(1/(COUNTA({base_col_letter}$1:{latest_col_letter}$1)-1))-1")
            analysis_ws[f'{growth_col_letter}{row_num}'] = cagr_formula

        # Q30: 空行（対応する行30が空行）
        # （analysis_wsの行30は空行なので何もしない）

        # Q31-Q33: 計算指標
        for idx, row_num in enumerate(yoy_rows_calc):
            source_cagr_row = [equity_row_num, equity_avg_row_num, total_assets_avg_row_num][idx]
            cagr_formula = (f"=({latest_col_letter}{source_cagr_row}/{base_col_letter}{source_cagr_row})"
                          f"^(1/(COUNTA({base_col_letter}$1:{latest_col_letter}$1)-1))-1")
            analysis_ws[f'{growth_col_letter}{row_num}'] = cagr_formula

        # Q34: 空行（対応する行34が空行）

        # Q35-Q41: ROE分析指標（check2_row_numは除外）
        source_rows_roe = [roe_calc_row_num, ros_row_num, tor_row_num, lrv_row_num,
                          check1_row_num, check2_row_num, roa_row_num]
        for idx, row_num in enumerate(yoy_rows_roe):
            source_cagr_row = source_rows_roe[idx]
            # check2_row_num（検算2）の場合は空欄
            if source_cagr_row == check2_row_num:
                continue
            cagr_formula = (f"=({latest_col_letter}{source_cagr_row}/{base_col_letter}{source_cagr_row})"
                          f"^(1/(COUNTA({base_col_letter}$1:{latest_col_letter}$1)-1))-1")
            analysis_ws[f'{growth_col_letter}{row_num}'] = cagr_formula

        # Q列の列幅を12に設定
        analysis_ws.column_dimensions[growth_col_letter].width = 12

        # Q24-Q41: 対前年増加率セクションの年平均増加率の表示形式を設定（パーセント）
        for row_num in yoy_rows_basic + yoy_rows_calc + yoy_rows_roe:
            analysis_ws[f'{growth_col_letter}{row_num}'].number_format = number_format_percent

    debug_log(f"ROE analysis sheet created: {analysis_sheet_name}")


def add_financial_analysis_sheets(workbook, debug_log=None):
    """
    財務分析シートを追加する（メイン関数）

    Args:
        workbook: openpyxlワークブック
        debug_log: デバッグログ関数（オプション）
    """
    # デバッグログ関数がない場合はダミー関数を使用
    if debug_log is None:
        def debug_log(msg):
            pass

    # 主要な経営指標等の推移（連結）シートを検索してROE分析シートを生成
    for sheet_name in workbook.sheetnames:
        if '主要な経営指標等の推移' in sheet_name and '連結' in sheet_name and '_' not in sheet_name:
            # "_"が含まれない = オリジナルシート（分析シートではない）
            try:
                create_roe_analysis_sheet(workbook, sheet_name, debug_log)
            except Exception as e:
                debug_log(f"Warning: Failed to create ROE analysis sheet for '{sheet_name}': {e}")
                # ROE分析シート生成に失敗してもメイン処理は継続
