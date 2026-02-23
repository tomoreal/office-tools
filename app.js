// 財務データ 横展開変換ロジック (JavaScript版)
const TARGET_SHEET_NAMES = [
    "連結貸借対照表",
    "連結損益計算書",
    "連結包括利益計算書",
    "連結損益（及び包括利益）計算書",
    "連結キャッシュ・フロー計算書",
    "連結株主資本等変動計算書"
];

function processFinancialCSV(csvText) {
    const lines = csvText.split(/\r?\n/);
    
    let dictData = {};
    let dictYears = new Set();
    let dictItemsOrder = {};
    let dictItemNames = {};
    
    let currentYear = "";
    let firstYear = "";
    let currentType = "";
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
            currentYear = col0.replace("現在", "").trim();
            dictYears.add(currentYear);
            if (firstYear === "") firstYear = currentYear;
            continue;
        }
        
        if (col0 === "表名称") {
            if (TARGET_SHEET_NAMES.includes(col1)) {
                currentType = col1;
                prevItem = "";
                hierarchyStack = [];
                if (!dictData[currentType]) {
                    dictData[currentType] = {};
                    dictItemsOrder[currentType] = [];
                    dictItemNames[currentType] = {};
                }
            }
            continue;
        }
        
        if (["企業名", "証券ｺｰﾄﾞ", "（百万円）"].includes(col0) || (col0.includes("/") && col0.includes("-"))) {
            continue;
        }
        
        if (col0 !== "" && col1 !== "" && col2 === "" && col3 === "") {
            if (TARGET_SHEET_NAMES.includes(col0)) {
                currentType = col0;
                prevItem = "";
                hierarchyStack = [];
                if (!dictData[currentType]) {
                    dictData[currentType] = {};
                    dictItemsOrder[currentType] = [];
                    dictItemNames[currentType] = {};
                }
                continue;
            }
        }
        
        if (col0 !== "" && currentType !== "") {
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
            
            dictItemNames[currentType][uniqueKey] = itemName;
            
            if (!dictItemsOrder[currentType].includes(uniqueKey)) {
                if (currentYear !== firstYear && dictItemsOrder[currentType].includes(prevItem)) {
                    const idx = dictItemsOrder[currentType].indexOf(prevItem);
                    dictItemsOrder[currentType].splice(idx + 1, 0, uniqueKey);
                } else {
                    dictItemsOrder[currentType].push(uniqueKey);
                }
            }
            
            prevItem = uniqueKey;
            
            if (!dictData[currentType][uniqueKey]) {
                dictData[currentType][uniqueKey] = {};
            }
            
            if (currentYear !== "" && amount !== "") {
                dictData[currentType][uniqueKey][currentYear] = amount;
            }
        }
    }
    
    return {
        dictData,
        dictYears: Array.from(dictYears).sort(),
        dictItemsOrder,
        dictItemNames
    };
}

function generateExcelWorkbook(parsedData) {
    const wb = XLSX.utils.book_new();
    const sortedYears = parsedData.dictYears;
    
    Object.keys(parsedData.dictData).forEach(typeName => {
        const items = parsedData.dictItemsOrder[typeName];
        if (items.length === 0) return;
        
        // 最大31文字かつ特殊文字を除外したシート名
        const safeTitle = typeName.replace(/[・ \/]/g, "").substring(0, 31);
        if (!safeTitle) return;
        
        const wsData = [];
        
        // ヘッダー行
        const headerRow = ["勘定科目", ...sortedYears];
        wsData.push(headerRow);
        
        // データ行
        items.forEach(uniqueKey => {
            const displayName = parsedData.dictItemNames[typeName][uniqueKey] || uniqueKey;
            const row = [displayName];
            
            sortedYears.forEach(year => {
                let val = parsedData.dictData[typeName][uniqueKey][year] || "";
                if (val !== "") {
                    // 数値に変換可能なものは数値にする
                    const numericVal = Number(val.replace(/,/g, ""));
                    if (!isNaN(numericVal) && val !== "-") {
                        val = numericVal;
                    }
                }
                row.push(val);
            });
            wsData.push(row);
        });
        
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, safeTitle);
    });
    
    return wb;
}

// UI要素のイベントハンドリング
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileNameDisplay = document.getElementById('fileName');
    const submitBtn = document.getElementById('submitBtn');
    const dropZone = document.getElementById('dropZone');

    fileInput.addEventListener('change', function() {
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
