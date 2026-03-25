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

    # ヘッダー行をコピー
    header_row = []
    for col in range(1, source_ws.max_column + 1):
        cell = source_ws.cell(1, col)
        header_row.append(cell.value)
    analysis_ws.append(header_row)

    # 列数を取得
    num_cols = len(header_row)

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

    # セクションヘッダー
    analysis_ws.append(['　連結経営指標等', 'jpcrp_cor_BusinessResultsOfGroupHeading'] + [''] * (num_cols - 2))

    # 基本指標（元シートから参照）
    def add_reference_row(label, english_name, source_row_num):
        """元シートのデータを参照する行を追加"""
        if source_row_num is None:
            return
        row_data = [label, english_name]
        for col in range(3, num_cols + 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            # 元シートのセルを参照する数式
            formula = f"=IF(ISBLANK('{source_sheet_name}'!{col_letter}{source_row_num}),\"\",'{source_sheet_name}'!{col_letter}{source_row_num})"
            row_data.append(formula)
        analysis_ws.append(row_data)

    add_reference_row('　　　売上高', sales_key, sales_row)
    add_reference_row('　　　親会社株主に帰属する当期純利益又は親会社株主に帰属する当期純損失（△）', profit_key, profit_row)
    add_reference_row('　　　純資産額', net_assets_key, net_assets_row)
    add_reference_row('　　　総資産額', total_assets_key, total_assets_row)
    add_reference_row('　　　自己資本比率', equity_ratio_key, equity_ratio_row)
    add_reference_row('　　　自己資本利益率', roe_key, roe_row)

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
