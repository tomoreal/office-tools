// app2.js

// csv-parse の簡易実装 (クオート対応)
function parseCSVLine(text) {
    const result = [];
    let curStr = '';
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        if (inQuotes) {
            if (char === '"') {
                if (text[i + 1] === '"') {
                    curStr += '"';
                    i++; // Skip next quote
                } else {
                    inQuotes = false;
                }
            } else {
                curStr += char;
            }
        } else {
            if (char === '"') {
                inQuotes = true;
            } else if (char === ',') {
                result.push(curStr);
                curStr = '';
            } else {
                curStr += char;
            }
        }
    }
    result.push(curStr);
    return result;
}

function processFinancialCSV(csvText) {
    const lines = csvText.split(/\r?\n/);
    let yearStandards = {}; // yearStandards[year] = "IFRS" | "日本基準"

    let scanYear = "";
    let scanInBSSection = false; // 貸借対照表セクション内かどうか

    lines.forEach(line => {
        if (!line.trim()) return;
        const row = parseCSVLine(line);
        if (row.length < 1) return;
        const col0 = row[0].trim();
        const col1 = row.length > 1 ? row[1].trim() : "";

        if (col0.includes("現在") && (col0.includes("/") || col0.includes("年"))) {
            // 年度文字列を正規化 (2024/3/31 -> 2024/03/31)
            let rawYear = col0.replace("現在", "").trim();
            scanYear = rawYear.replace(/(\d+)\/(\d+)\/(\d+)/, (_match, y, m1, d) => {
                return `${y}/${m1.padStart(2, '0')}/${d.padStart(2, '0')}`;
            });
            scanInBSSection = false;
            if (!yearStandards[scanYear]) {
                yearStandards[scanYear] = "日本基準"; // 基本は日本基準とする
            }
        } else if (col0 === "表名称") {
            const rawType = col1;
            // 貸借対照表系の表名（連結貸借対照表または連結財政状態計算書）
            if (rawType === "連結貸借対照表" || rawType === "連結財政状態計算書") {
                scanInBSSection = true;
                // 「連結財政状態計算書」という表名自体がIFRSの証拠
                if (scanYear && rawType === "連結財政状態計算書") {
                    yearStandards[scanYear] = "IFRS";
                }
            } else {
                scanInBSSection = false;
            }
        } else if (col0 === "連結財政状態計算書" || col0 === "連結貸借対照表") {
            // データ行の中にシート名が出現する場合（表名称行がない場合）
            scanInBSSection = true;
            if (scanYear && col0 === "連結財政状態計算書") {
                yearStandards[scanYear] = "IFRS";
            }
        } else if (scanYear && scanInBSSection && col0 !== "") {
            // 貸借対照表セクション内で「非流動資産」「非流動負債」が出現したらIFRS
            if (col0.includes("非流動資産") || col0.includes("非流動負債")) {
                yearStandards[scanYear] = "IFRS";
            }
        }
    });

    return yearStandards;
}

/**
 * 項目名の正規化 (横並びの不一致を解消するため)
 */
function normalizeKey(str) {
    if (!str) return "";
    return str.normalize('NFKC')
        .replace(/\s+/g, '') // 空白除去
        .replace(/[・\.．、，]/g, '') // 記号の揺れを除去
        .replace(/[（\(\[]?△は(?:損失|減少|増加|利益)[）\)\]]?/g, "") // (△は減少) などを除去
        .replace(/[（\(\)）]/g, (m) => ({ '（': '(', '）': ')', '(': '(', ')': ')' }[m]));
}

const SHEET_MAPPING = {
    "連結貸借対照表": "連結貸借対照表",
    "連結財政状態計算書": "連結貸借対照表",
    "連結損益計算書": "連結損益計算書",
    "連結包括利益計算書": "連結損益計算書",
    "連結損益（及び包括利益）計算書": "連結損益計算書",
    "連結キャッシュ・フロー計算書": "連結キャッシュ・フロー計算書",
    "連結株主資本等変動計算書": "連結株主資本等変動計算書"
};

const TARGET_SHEET_NAMES = Object.keys(SHEET_MAPPING);

/**
 * CSVの行を縦持ちデータのオブジェクト配列にパースする
 */
function parseFinancialData(csvText, yearStandards) {
    const lines = csvText.split(/\r?\n/);
    const records = [];

    let currentYear = "";
    let currentBaseType = "";
    let accountNumberOriginal = 0;

    lines.forEach(line => {
        if (!line.trim()) return;
        const row = parseCSVLine(line);
        if (row.length < 1) return;

        let rawCol0 = row[0];
        let col0 = rawCol0.trim();
        const col1 = row.length > 1 ? row[1].trim() : "";
        const col2 = row.length > 2 ? row[2].trim() : "";
        const col3 = row.length > 3 ? row[3].trim() : "";

        if (col0.includes("現在") && (col0.includes("/") || col0.includes("年"))) {
            let rawYear = col0.replace("現在", "").trim();
            currentYear = rawYear.replace(/(\d+)\/(\d+)\/(\d+)/, (m, y, m1, d) => {
                return `${y}/${m1.padStart(2, '0')}/${d.padStart(2, '0')}`;
            });
            return;
        }

        if (col0 === "表名称") {
            const rawType = col1;
            if (TARGET_SHEET_NAMES.includes(rawType)) {
                currentBaseType = SHEET_MAPPING[rawType] || rawType;
                accountNumberOriginal = 0;
            } else {
                currentBaseType = "";
            }
            return;
        }

        if (TARGET_SHEET_NAMES.includes(col0)) {
            const rawType = col0;
            currentBaseType = SHEET_MAPPING[rawType] || rawType;
            accountNumberOriginal = 0;
            return;
        }

        // 不要行のスキップ (「（百万円）」などを除外)
        if (["企業名", "証券ｺｰﾄﾞ", "（百万円）", "(百万円)"].includes(col0) || (col0.includes("/") && col0.includes("-"))) {
            return;
        }

        // データの格納
        if (col0 !== "" && currentBaseType !== "" && currentYear !== "") {
            const standard = yearStandards[currentYear] || "日本基準";
            let amount = "";

            const isNumericAmount = (val) => {
                const cleaned = val.replace(/-/g, "").replace(/,/g, "");
                return val === "-" || (!isNaN(cleaned) && cleaned.length > 0);
            };

            if (col2 !== "" && isNumericAmount(col2)) {
                amount = col2;
            } else if (col3 !== "" && isNumericAmount(col3)) {
                amount = col3;
            }

            const normalName = normalizeKey(rawCol0.trim());
            accountNumberOriginal++;

            // 表示用名（元のインデント等を含む、最初に見つけた名称を維持するために使用）
            const displayName = rawCol0;

            records.push({
                year: currentYear,
                accounting_standard: standard,
                financial_statement_type: currentBaseType,
                account_name_original: rawCol0,
                account_name_display: displayName,
                account_name_normal: normalName,
                account_number_original: accountNumberOriginal,
                value: amount
            });
        }
    });

    return records;
}

/**
 * 縦持ち配列から、会計基準とシートごとの「マスター配列(正しい並べ順)」を構築する
 */
// 戻り値: { masterLists: { [ standard ]: { [ type ]: [ name_normal, ... ] } }, displayNames: { [ name_normal ]: name_display } }
function buildMasterLists(records) {
    const grouped = {};
    const displayNames = {}; // normal_name -> display_name

    // 1. 会計基準 > 財務諸表 > 年度 ごとにグルーピングしつつ、表示名を記録
    records.forEach(rc => {
        const st = rc.accounting_standard;
        const ft = rc.financial_statement_type;
        const yr = rc.year;

        if (!grouped[st]) grouped[st] = {};
        if (!grouped[st][ft]) grouped[st][ft] = {};
        if (!grouped[st][ft][yr]) grouped[st][ft][yr] = [];

        grouped[st][ft][yr].push(rc);

        // 表示名の保存（最初に見つかったもの、できればインデントありのものを優先）
        if (!displayNames[rc.account_name_normal] || rc.account_name_display.startsWith(" ")) {
            displayNames[rc.account_name_normal] = rc.account_name_display;
        }
    });

    const masterLists = {}; // masterLists[standard][type] = [ name_normal, ... ]

    // 2. グループごとにマスター配列を逆算して構築
    for (const st in grouped) {
        masterLists[st] = {};
        for (const ft in grouped[st]) {
            let masterList = [];

            // 最新の年度から過去へと並べ替え (降順)
            const presentYears = Object.keys(grouped[st][ft]).sort((a, b) => b.localeCompare(a));
            if (presentYears.length === 0) continue;

            const latestYear = presentYears[0];
            const latestItems = grouped[st][ft][latestYear];

            // ① 最新年度の科目を順番通りにマスター配列へ
            latestItems.forEach(item => {
                if (!masterList.includes(item.account_name_normal)) {
                    masterList.push(item.account_name_normal);
                }
            });

            // ② 過去の年度を順番に辿って未登録科目を挿入 (account_name_exist == 0 の処理)
            for (let i = 1; i < presentYears.length; i++) {
                const year = presentYears[i];
                const items = grouped[st][ft][year].map(x => x.account_name_normal);

                for (let j = 0; j < items.length; j++) {
                    const itemName = items[j];

                    if (masterList.includes(itemName)) {
                        continue;
                    }

                    // その0の行から下に、存在する(1)行を探す
                    let nextExistingItemIndex = -1;
                    for (let k = j + 1; k < items.length; k++) {
                        if (masterList.includes(items[k])) {
                            nextExistingItemIndex = k;
                            break;
                        }
                    }

                    if (nextExistingItemIndex !== -1) {
                        const nextItemName = items[nextExistingItemIndex];
                        const insertIndex = masterList.indexOf(nextItemName);
                        masterList.splice(insertIndex, 0, itemName);
                    } else {
                        masterList.push(itemName);
                    }
                }
            }
            masterLists[st][ft] = masterList;
        }
    }

    return { masterLists, displayNames };
}

/**
 * 構築したマスターリストと縦持ちレコードからExcelワークブックを生成する
 */
function generateExcelWorkbook(records, buildResult) {
    const wb = XLSX.utils.book_new();
    const { masterLists, displayNames } = buildResult;

    // 全年度のリストを取得し、昇順（古い順）に並べる（横軸表示用）
    const allYears = Array.from(new Set(records.map(r => r.year))).sort();

    // { [standard]: { [type]: { [year]: { [name_normal]: value } } } }
    const pivotData = {};
    records.forEach(rc => {
        const st = rc.accounting_standard;
        const ft = rc.financial_statement_type;
        const yr = rc.year;
        const nm = rc.account_name_normal;

        if (!pivotData[st]) pivotData[st] = {};
        if (!pivotData[st][ft]) pivotData[st][ft] = {};
        if (!pivotData[st][ft][yr]) pivotData[st][ft][yr] = {};

        pivotData[st][ft][yr][nm] = rc.value;
    });

    for (const st in masterLists) {
        for (const ft in masterLists[st]) {
            const masterList = masterLists[st][ft]; // このシートの科目の正しい並び順

            // このグループに含まれる年度のみを抽出（表の列方向）
            const targetYears = allYears.filter(yr => pivotData[st][ft] && pivotData[st][ft][yr]);
            if (targetYears.length === 0) continue;

            // シート名の決定
            let suffix = (st === "IFRS") ? "(IFRS)" : "(日本基準)";
            let sheetTitle = ft + suffix;

            // 特殊対応: IFRSの貸借対照表は「連結財政状態計算書(IFRS)」とする
            if (ft === "連結貸借対照表" && st === "IFRS") {
                sheetTitle = "連結財政状態計算書(IFRS)";
            }

            const safeTitle = sheetTitle.replace(/[・ \/]/g, "").substring(0, 31);

            // Excel表データの作成 (2次元配列)
            const wsData = [];
            const headerRow = ["勘定科目", ...targetYears];
            wsData.push(headerRow);

            const colWidths = new Array(headerRow.length).fill(0);
            const getDisplayWidth = (str) => {
                if (str === null || str === undefined) return 0;
                const s = String(str);
                let len = 0;
                for (let i = 0; i < s.length; i++) {
                    len += (s.charCodeAt(i) > 255) ? 2 : 1;
                }
                return len;
            };

            headerRow.forEach((h, i) => {
                colWidths[i] = getDisplayWidth(h);
            });

            // マスターリストの順番に従って行を出力
            masterList.forEach(normalName => {
                const displayName = displayNames[normalName] || normalName;

                const aColWidth = getDisplayWidth(displayName);
                if (aColWidth > colWidths[0]) colWidths[0] = aColWidth;

                const row = [displayName];

                targetYears.forEach((year, idx) => {
                    let val = pivotData[st][ft][year]?.[normalName] || "";
                    if (val !== "") {
                        const numericVal = Number(val.replace(/,/g, ""));
                        if (!isNaN(numericVal) && val !== "-") {
                            val = numericVal;
                            const valWidth = getDisplayWidth(val.toLocaleString('en-US'));
                            if (valWidth > colWidths[idx + 1]) colWidths[idx + 1] = valWidth;
                        } else {
                            const valWidth = getDisplayWidth(val);
                            if (valWidth > colWidths[idx + 1]) colWidths[idx + 1] = valWidth;
                        }
                    }
                    row.push(val);
                });

                wsData.push(row);
            });

            const ws = XLSX.utils.aoa_to_sheet(wsData);
            ws['!cols'] = colWidths.map(w => ({ wch: w + 2 }));

            for (let cellAddress in ws) {
                if (cellAddress.startsWith("!")) continue;
                if (!ws[cellAddress].s) ws[cellAddress].s = {};
                if (!ws[cellAddress].s.font) ws[cellAddress].s.font = {};
                ws[cellAddress].s.font.name = "ＭＳ 明朝";
                ws[cellAddress].s.font.sz = 10;
                if (ws[cellAddress].t === "n") {
                    ws[cellAddress].z = '#,##0_ ;[Red]\\-#,##0\\ ';
                }
            }

            XLSX.utils.book_append_sheet(wb, ws, safeTitle);
        }
    }

    return wb;
}

// UI要素のイベントハンドリング
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileNameDisplay = document.getElementById('fileName');
    const submitBtn = document.getElementById('submitBtn');
    const dropZone = document.getElementById('dropZone');

    fileInput.addEventListener('change', function () {
        if (this.files && this.files.length > 0) {
            fileNameDisplay.textContent = this.files[0].name;
            submitBtn.disabled = false;
        } else {
            fileNameDisplay.textContent = '';
            submitBtn.disabled = true;
        }
    });

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, e => {
            e.preventDefault();
            e.stopPropagation();
        }, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
    });

    dropZone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            if (files[0].name.toLowerCase().endsWith('.csv')) {
                fileInput.files = files;
                fileNameDisplay.textContent = files[0].name;
                submitBtn.disabled = false;
            } else {
                alert('CSVファイルのみを選択してください。');
                submitBtn.disabled = true;
            }
        }
    }, false);

    submitBtn.addEventListener('click', async () => {
        if (!fileInput.files || fileInput.files.length === 0) return;

        const file = fileInput.files[0];
        const outFileName = file.name.replace(/\.[^/.]+$/, "") + '_横展開.xlsx';

        submitBtn.classList.add('loading');
        submitBtn.disabled = true;

        try {
            // Shift_JIS (cp932) としてファイルを読み込む
            const buffer = await file.arrayBuffer();
            const decoder = new TextDecoder('shift-jis');
            const csvText = decoder.decode(buffer);

            // 1. 各年度ごとの会計基準判定
            const standards = processFinancialCSV(csvText);

            // 2. CSV全行を縦持ちデータのオブジェクトとしてパース
            const records = parseFinancialData(csvText, standards);

            // 3. アレイベース（JS配列）での勘定科目並び順ロジックの実行
            const buildResult = buildMasterLists(records);

            // 4. Excelファイルの生成
            const wb = generateExcelWorkbook(records, buildResult);

            // Excelファイルのダウンロード
            XLSX.writeFile(wb, outFileName);

            showFlashMessage('success', '変換が完了し、ダウンロードが開始されました。');

        } catch (error) {
            console.error(error);
            showFlashMessage('error', '変換中にエラーが発生しました: ' + error.message);
        } finally {
            submitBtn.classList.remove('loading');
            submitBtn.disabled = false;
        }
    });
});

function showFlashMessage(type, message) {
    const container = document.querySelector('.container');
    const oldFlash = document.querySelector('.flash-messages');
    if (oldFlash) oldFlash.remove();

    const flashContainer = document.createElement('div');
    flashContainer.className = 'flash-messages';

    const flashMsg = document.createElement('div');
    flashMsg.className = `flash ${type}`;
    flashMsg.textContent = message;

    flashContainer.appendChild(flashMsg);

    // ヘッダーの直後に挿入
    const header = document.querySelector('.header');
    header.parentNode.insertBefore(flashContainer, header.nextSibling);

    // 5秒後に消える
    setTimeout(() => flashMsg.remove(), 5000);
}
