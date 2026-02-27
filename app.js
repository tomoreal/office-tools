// 財務データ 横展開変換ロジック (JavaScript版) v1.2
const TARGET_SHEET_NAMES = [
    "連結貸借対照表",
    "連結財政状態計算書", // IFRS BS
    "連結損益計算書",
    "連結包括利益計算書",
    "連結損益（及び包括利益）計算書",
    "連結キャッシュ・フロー計算書",
    "連結株主資本等変動計算書"
];

const SHEET_MAPPING = {
    "連結貸借対照表": "連結貸借対照表",
    "連結財政状態計算書": "連結財政状態計算書",
    "連結損益計算書": "連結損益計算書",
    "連結包括利益計算書": "連結損益計算書", // 損益系は統合
    "連結損益（及び包括利益）計算書": "連結損益計算書", // 損益系は統合
    "連結キャッシュ・フロー計算書": "連結キャッシュ・フロー計算書",
    "連結株主資本等変動計算書": "連結株主資本等変動計算書"
};

function processFinancialCSV(csvText) {
    const lines = csvText.split(/\r?\n/);

    let dictData = {}; // dictData[baseType][uniqueKey][year]
    let dictYears = new Set();
    let dictItemsOrder = {}; // dictItemsOrder[baseType] = []
    let dictItemNames = {}; // dictItemNames[baseType][uniqueKey] = itemName
    let yearStandards = {}; // yearStandards[year] = "IFRS" | "J-GAAP"

    let currentYear = "";
    let firstYear = "";
    let currentBaseType = "";
    let prevItem = "";
    let hierarchyStack = [];

    // csv-parse の簡易実装 (クオート対応)
    const parseCSVLine = (text) => {
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
    };

    for (let rowIdx = 0; rowIdx < lines.length; rowIdx++) {
        const line = lines[rowIdx];
        if (!line.trim()) continue;

        const row = parseCSVLine(line);
        if (row.length < 1) continue;

        const rawCol0 = row[0];
        const col0 = rawCol0.trim();
        const col1 = row.length > 1 ? row[1].trim() : "";
        const col2 = row.length > 2 ? row[2].trim() : "";
        const col3 = row.length > 3 ? row[3].trim() : "";

        if (col0.includes("現在") && (col0.includes("/") || col0.includes("年"))) {
            // 年度文字列の正規化 (2024/3/31 -> 2024/03/31)
            let rawYear = col0.replace("現在", "").trim();
            currentYear = rawYear.replace(/(\d+)\/(\d+)\/(\d+)/, (m, y, m1, d) => {
                return `${y}/${m1.padStart(2, '0')}/${d.padStart(2, '0')}`;
            });
            dictYears.add(currentYear);
            if (firstYear === "") firstYear = currentYear;
            continue;
        }

        if (col0 === "表名称") {
            const rawType = col1;
            if (TARGET_SHEET_NAMES.includes(rawType)) {
                currentBaseType = SHEET_MAPPING[rawType] || rawType;

                // IFRS判定: IFRS優先で設定（currentYearが空でない場合のみ）
                if (currentYear) {
                    if (rawType === "連結財政状態計算書") {
                        yearStandards[currentYear] = "IFRS";
                    } else if (rawType === "連結貸借対照表") {
                        if (yearStandards[currentYear] !== "IFRS") {
                            yearStandards[currentYear] = "J-GAAP";
                        }
                    }
                }

                prevItem = "";
                hierarchyStack = [];
                if (!dictData[currentBaseType]) {
                    dictData[currentBaseType] = {};
                    dictItemsOrder[currentBaseType] = [];
                    dictItemNames[currentBaseType] = {};
                }
            } else {
                currentBaseType = "";
            }
            continue;
        }

        if (["企業名", "証券ｺｰﾄﾞ", "（百万円）"].includes(col0) || (col0.includes("/") && col0.includes("-"))) {
            continue;
        }

        if (col0 !== "" && col1 !== "" && col2 === "" && col3 === "") {
            const rawType = col0;
            if (TARGET_SHEET_NAMES.includes(rawType)) {
                currentBaseType = SHEET_MAPPING[rawType] || rawType;

                if (currentYear) {
                    if (rawType === "連結財政状態計算書") {
                        yearStandards[currentYear] = "IFRS";
                    } else if (rawType === "連結貸借対照表") {
                        if (yearStandards[currentYear] !== "IFRS") {
                            yearStandards[currentYear] = "J-GAAP";
                        }
                    }
                }

                prevItem = "";
                hierarchyStack = [];
                if (!dictData[currentBaseType]) {
                    dictData[currentBaseType] = {};
                    dictItemsOrder[currentBaseType] = [];
                    dictItemNames[currentBaseType] = {};
                }
                continue;
            }
        }

        if (col0 !== "" && currentBaseType !== "") {
            // IFRS特有の項目名が含まれる場合はIFRS判定にする (連結貸借対照表であっても)
            if (currentYear && (col0.includes("非流動資産") || col0.includes("非流動負債"))) {
                yearStandards[currentYear] = "IFRS";
            }

            const itemName = rawCol0.trimEnd();
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

            // 全角・半角スペースとタブによるインデントレベルの計算
            const strippedLen = rawCol0.replace(/^[ \t　]+/, '').length;
            const indentLevel = rawCol0.length - strippedLen;

            while (hierarchyStack.length > 0 && hierarchyStack[hierarchyStack.length - 1][0] >= indentLevel) {
                hierarchyStack.pop();
            }
            hierarchyStack.push([indentLevel, col0]);

            const uniqueKey = hierarchyStack.map(x => x[1]).join("::");

            dictItemNames[currentBaseType][uniqueKey] = itemName;

            if (!dictItemsOrder[currentBaseType].includes(uniqueKey)) {
                if (currentYear !== firstYear && dictItemsOrder[currentBaseType].includes(prevItem)) {
                    const idx = dictItemsOrder[currentBaseType].indexOf(prevItem);
                    dictItemsOrder[currentBaseType].splice(idx + 1, 0, uniqueKey);
                } else {
                    dictItemsOrder[currentBaseType].push(uniqueKey);
                }
            }

            prevItem = uniqueKey;

            if (!dictData[currentBaseType][uniqueKey]) {
                dictData[currentBaseType][uniqueKey] = {};
            }

            if (currentYear !== "" && amount !== "") {
                dictData[currentBaseType][uniqueKey][currentYear] = amount;
            }
        }
    }

    // 最終的な年度ごとの基準を確定させる（一度IFRSになった後は、以降の年度もすべてIFRS）
    const sortedYears = Array.from(dictYears).filter(y => y !== "").sort();
    let reachedIFRS = false;
    sortedYears.forEach(year => {
        if (yearStandards[year] === "IFRS") {
            reachedIFRS = true;
        }
        if (reachedIFRS) {
            yearStandards[year] = "IFRS";
        } else {
            if (yearStandards[year] !== "IFRS") {
                yearStandards[year] = "J-GAAP";
            }
        }
    });

    return {
        dictData,
        dictYears: sortedYears,
        dictItemsOrder,
        dictItemNames,
        yearStandards
    };
}

function generateExcelWorkbook(parsedData) {
    const wb = XLSX.utils.book_new();
    const sortedYears = parsedData.dictYears;
    const yearStandards = parsedData.yearStandards;

    Object.keys(parsedData.dictData).forEach(baseType => {
        // 日本基準とIFRSでデータを分ける
        const standards = ["J-GAAP", "IFRS"];

        standards.forEach(std => {
            const targetYears = sortedYears.filter(y => yearStandards[y] === std);
            if (targetYears.length === 0) return;

            const items = parsedData.dictItemsOrder[baseType];
            // この基準（std）において、少なくとも1つの年度で値が存在する項目のみを抽出
            const validItems = items.filter(uniqueKey => {
                return targetYears.some(year => parsedData.dictData[baseType][uniqueKey][year] !== undefined);
            });

            if (validItems.length === 0) return;

            // シート名の決定
            let suffix = (std === "IFRS") ? "(IFRS)" : "(日本基準)";
            let sheetTitle = baseType + suffix;

            // 最大31文字かつ特殊文字を除外したシート名
            const safeTitle = sheetTitle.replace(/[・ \/]/g, "").substring(0, 31);

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

            validItems.forEach(uniqueKey => {
                const displayName = parsedData.dictItemNames[baseType][uniqueKey] || uniqueKey;
                const aColWidth = getDisplayWidth(displayName);
                if (aColWidth > colWidths[0]) colWidths[0] = aColWidth;

                const row = [displayName];
                targetYears.forEach((year, idx) => {
                    let val = parsedData.dictData[baseType][uniqueKey][year] || "";
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
        });
    });

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

            // 変換処理
            const parsedData = processFinancialCSV(csvText);
            const wb = generateExcelWorkbook(parsedData);

            // Excelファイルのダウンロード
            XLSX.writeFile(wb, outFileName);

            // 成功メッセージ（Flash代わり）
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
