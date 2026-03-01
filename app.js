// 財務データ 横展開変換ロジック (JavaScript版) v1.6
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
    "連結財政状態計算書": "連結貸借対照表", // 貸借対照表系は内部で「連結貸借対照表」に統合
    "連結損益計算書": "連結損益計算書",
    "連結包括利益計算書": "連結損益計算書", // 損益系は統合
    "連結損益（及び包括利益）計算書": "連結損益計算書", // 損益系は統合
    "連結キャッシュ・フロー計算書": "連結キャッシュ・フロー計算書",
    "連結株主資本等変動計算書": "連結株主資本等変動計算書"
};

/**
 * 項目名の正規化 (横並びの不一致を解消するため)
 */
function normalizeKey(str) {
    if (!str) return "";
    return str.normalize('NFKC')
        .replace(/\s+/g, '') // 空白除去
        .replace(/[・\.．、，]/g, '') // 記号の揺れを除去
        .replace(/[（\(\)）]/g, (m) => ({ '（': '(', '）': ')', '(': '(', ')': ')' }[m]));
}

function processFinancialCSV(csvText) {
    const lines = csvText.split(/\r?\n/);

    let dictData = {}; // dictData[baseType][uniqueKey][year]
    let dictYears = new Set();
    let dictItemsOrder = {}; // dictItemsOrder[baseType] = []
    let dictItemNames = {}; // dictItemNames[baseType][uniqueKey] = itemName
    let dictIsHeader = {}; // dictIsHeader[baseType][uniqueKey] = Set([year1, year2, ...])
    let dictItemFirstSeen = {}; // dictItemFirstSeen[baseType][uniqueKey] = {yearIndex, rowIndex}
    let yearStandards = {}; // yearStandards[year] = "IFRS" | "J-GAAP"

    let currentYear = "";
    let yearIndex = -1; // 年度の出現順序
    let currentBaseType = "";
    let hierarchyStack = []; // [ {level, name, uniqueKey} ]
    let currentLandmarks = {
        major: "", // 資産, 負債, 資本
        sub: ""    // 流動, 非流動
    };
    let currentPLSubsection = ""; // 損益計算書の中間セクション（非階層化期間用）

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

    // 事前スキャン1：年度ごとの基準（IFRSかどうか）を特定する
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

    // 事前スキャン2：年度×会計基準ごとに階層化の有無を検出
    let yearStandardHierarchy = {}; // {year_standard: {hasIndent: count, noIndent: count}}
    scanYear = "";
    scanBaseType = "";
    let scanStandard = "";

    lines.forEach(line => {
        if (!line.trim()) return;
        const row = parseCSVLine(line);
        if (row.length < 1) return;
        const rawCol0 = row[0];
        const col0 = rawCol0.trim();
        const col1 = row.length > 1 ? row[1].trim() : "";
        const col2 = row.length > 2 ? row[2].trim() : "";
        const col3 = row.length > 3 ? row[3].trim() : "";

        if (col0.includes("現在") && (col0.includes("/") || col0.includes("年"))) {
            let rawYear = col0.replace("現在", "").trim();
            scanYear = rawYear.replace(/(\d+)\/(\d+)\/(\d+)/, (_match, y, m1, d) => {
                return `${y}/${m1.padStart(2, '0')}/${d.padStart(2, '0')}`;
            });
            scanStandard = yearStandards[scanYear] || "J-GAAP";
            scanBaseType = "";
        } else if (col0 === "表名称") {
            const rawType = row.length > 1 ? row[1].trim() : "";
            if (TARGET_SHEET_NAMES.includes(rawType)) {
                scanBaseType = SHEET_MAPPING[rawType] || rawType;
            } else {
                scanBaseType = "";
            }
        } else if (TARGET_SHEET_NAMES.includes(col0)) {
            // データ行の中にシート名が出現する場合
            if (col1 === "" || col1.includes("Consolidated") || col1.includes("Statement")) {
                scanBaseType = SHEET_MAPPING[col0] || col0;
            }
        } else if (scanYear && ["連結貸借対照表", "連結損益計算書"].includes(scanBaseType) && col0 !== "") {
            // 貸借対照表と損益計算書の項目をチェック
            const key = `${scanYear}_${scanStandard}`;
            if (!yearStandardHierarchy[key]) {
                yearStandardHierarchy[key] = { hasIndent: 0, noIndent: 0, items: [] };
            }

            // 先頭スペースの有無をチェック（全角・半角スペース、タブ）
            const hasLeadingSpace = /^[ \t　]+/.test(rawCol0);

            // より広範な項目をチェック対象とする（表名やメタ情報は除外）
            const isMetaData = col0 === "連結貸借対照表" || col0 === "連結財政状態計算書" ||
                col0.includes("Consolidated") || col0.includes("（百万円）");
            const isLikelyDataItem = !isMetaData && col0.length > 2;

            if (isLikelyDataItem) {
                yearStandardHierarchy[key].items.push({ name: col0, hasIndent: hasLeadingSpace });
                if (hasLeadingSpace) {
                    yearStandardHierarchy[key].hasIndent++;
                } else {
                    yearStandardHierarchy[key].noIndent++;
                }
            }
        }
    });

    // 階層化されていない期間を特定
    let nonHierarchicalPeriods = new Set();
    Object.keys(yearStandardHierarchy).forEach(key => {
        const stats = yearStandardHierarchy[key];
        const total = stats.hasIndent + stats.noIndent;
        if (total > 0) {
            const indentRatio = stats.hasIndent / total;
            // インデント率が30%未満の場合、階層化されていないと判定
            if (indentRatio < 0.3) {
                nonHierarchicalPeriods.add(key);
                console.log(`[検出] 階層化なし: ${key}, インデント率=${(indentRatio * 100).toFixed(1)}% (${stats.hasIndent}/${total})`);
            }
        }
    });

    for (let rowIdx = 0; rowIdx < lines.length; rowIdx++) {
        const line = lines[rowIdx];
        if (!line.trim()) continue;

        const row = parseCSVLine(line);
        if (row.length < 1) continue;

        let rawCol0 = row[0];
        let col0 = rawCol0.trim();
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
            yearIndex++;
            continue;
        }

        if (col0 === "表名称") {
            const rawType = col1;
            if (TARGET_SHEET_NAMES.includes(rawType)) {
                currentBaseType = SHEET_MAPPING[rawType] || rawType;

                // IFRS判定は事前スキャンで完了しているので、ここでは上書きしない
                // 「連結財政状態計算書」が出現した場合のみ、確実にIFRSとマーク
                if (currentYear && rawType === "連結財政状態計算書") {
                    yearStandards[currentYear] = "IFRS";
                }

                hierarchyStack = [];
                currentLandmarks = { major: "", sub: "" };
                currentPLSubsection = ""; // 新しいシートに入ったらリセット
                if (!dictData[currentBaseType]) {
                    dictData[currentBaseType] = {};
                    dictItemsOrder[currentBaseType] = [];
                    dictItemNames[currentBaseType] = {};
                    dictIsHeader[currentBaseType] = {};
                    dictItemFirstSeen[currentBaseType] = {};
                }
            } else {
                currentBaseType = "";
            }
            continue;
        }

        if (["企業名", "証券ｺｰﾄﾞ", "（百万円）"].includes(col0) || (col0.includes("/") && col0.includes("-"))) {
            continue;
        }

        if (TARGET_SHEET_NAMES.includes(col0)) {
            const rawType = col0;
            currentBaseType = SHEET_MAPPING[rawType] || rawType;

            // IFRS判定は事前スキャンで完了しているので、ここでは上書きしない
            // 「連結財政状態計算書」が出現した場合のみ、確実にIFRSとマーク
            if (currentYear && rawType === "連結財政状態計算書") {
                yearStandards[currentYear] = "IFRS";
            }

            hierarchyStack = [];
            currentLandmarks = { major: "", sub: "" };
            currentPLSubsection = ""; // 新しいシートに入ったらリセット
            if (!dictData[currentBaseType]) {
                dictData[currentBaseType] = {};
                dictItemsOrder[currentBaseType] = [];
                dictItemNames[currentBaseType] = {};
                dictIsHeader[currentBaseType] = {};
                dictItemFirstSeen[currentBaseType] = {};
            }
            continue;
        }

        if (col0 !== "" && currentBaseType !== "") {
            // IFRS判定は事前スキャンで完了済み（ここでは追加の判定はしない）

            let itemName = rawCol0.trim(); // 前後の空白を完全に除去

            // J-GAAP等の「（内訳）」ラベルを「包括利益の帰属」に統一
            if (itemName === "（内訳）" || itemName === "(内訳)") {
                col0 = "包括利益の帰属";
                itemName = "包括利益の帰属";
                rawCol0 = rawCol0.replace(/（内訳）|\(内訳\)/, "包括利益の帰属");
            }

            // 古いIFRSの「持分法による投資損益（△は損失）」を統一
            if (col0.includes("持分法による投資損益") && col0.includes("△は損失")) {
                col0 = "持分法による投資損益";
                itemName = "持分法による投資損益";
                rawCol0 = "持分法による投資損益";
            }
            // col0 は既に先頭で定義されている（rawCol0.trim()済）
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

            // 現在の年度×会計基準が階層化されていないかチェック
            const currentStandard = yearStandards[currentYear] || "J-GAAP";
            const periodKey = `${currentYear}_${currentStandard}`;
            const isNonHierarchical = nonHierarchicalPeriods.has(periodKey);

            // 全角・半角スペースとタブによるインデントレベルの計算
            const strippedLen = rawCol0.replace(/^[ \t　]+/, '').length;
            let indentLevel = rawCol0.length - strippedLen;

            // 階層化されていない期間の場合、インデントレベルを推測
            // （ランドマークやキーワードから階層を推論）
            if (isNonHierarchical && indentLevel === 0) {
                const nName = normalizeKey(col0);
                // レベル1: 大セクション（資産、負債、資本）
                if (nName.includes("資産合計") || nName === "資産" || nName.includes("資産の部") ||
                    nName.includes("負債合計") || nName.includes("負債の部") || nName === "負債" ||
                    nName.includes("資本合計") || nName.includes("純資産合計") || nName === "資本" ||
                    nName === "純資産" || nName === "株主資本" || nName.includes("負債及び資本")) {
                    indentLevel = 1;
                }
                // レベル2: 中セクション（流動資産、非流動資産、流動負債、非流動負債）
                else if (nName === "流動資産" || nName === "非流動資産" || nName === "固定資産" ||
                    nName === "流動負債" || nName === "非流動負債" || nName === "固定負債") {
                    indentLevel = 2;
                }
                // レベル3: 詳細項目（現金及び現金同等物、など）と合計行
                else {
                    indentLevel = 3;
                }
            }

            const isSectionHeader = amount === "";

            // デバッグ: 「純損益に振り替えられる」を含む行を全て記録
            if (col0.includes("純損益に振り替えられる")) {
                console.log(`[デバッグ/純損益] year=${currentYear}, baseType=${currentBaseType}, col0="${col0}", amount="${amount}", isSectionHeader=${isSectionHeader}, isNonHierarchical=${isNonHierarchical}`);
            }

            // ランドマーク（主要セクション）の特定
            const nName = normalizeKey(col0);
            let landmarkChanged = false;
            const prevMajor = currentLandmarks.major;

            if (nName.includes("資産合計") || nName === "資産" || nName.includes("資産の部")) {
                currentLandmarks.major = "資産";
                currentLandmarks.sub = "";
            } else if (nName.includes("負債") && (nName.includes("資本") || nName.includes("純資産"))) {
                currentLandmarks.major = "負債及び資本";
                currentLandmarks.sub = "";
            } else if (nName.includes("負債合計") || nName.includes("負債の部") || nName === "負債") {
                currentLandmarks.major = "負債";
                currentLandmarks.sub = "";
            } else if (nName.includes("資本合計") || nName.includes("純資産合計") || nName.includes("資本の部") || nName === "資本" || nName === "純資産" || nName === "株主資本") {
                currentLandmarks.major = "資本";
                currentLandmarks.sub = "";
            }

            if (currentLandmarks.major !== prevMajor) {
                landmarkChanged = true;
                hierarchyStack = []; // メジャーセクションが変わったらスタックをリセット
            }

            // サブセクションの特定（実際のセクション名を使用）
            if (nName === "流動資産" || nName === "流動負債") {
                if (currentLandmarks.sub !== nName) landmarkChanged = true;
                currentLandmarks.sub = nName;
                // 資産か負債か未確定の場合は、文脈から推論
                if (!currentLandmarks.major) {
                    currentLandmarks.major = nName.includes("資産") ? "資産" : "負債";
                }
            } else if (nName === "非流動資産" || nName === "固定資産" || nName === "非流動負債" || nName === "固定負債") {
                if (currentLandmarks.sub !== nName) landmarkChanged = true;
                currentLandmarks.sub = nName;
                if (!currentLandmarks.major) {
                    currentLandmarks.major = (nName.includes("資産") || nName.includes("固定資産")) ? "資産" : "負債";
                }
            }

            // 損益計算書の中間セクション検出（非階層化期間用）
            if (isNonHierarchical && currentBaseType === "連結損益計算書" && isSectionHeader) {
                console.log(`[PL中間セクション検出] col0="${col0}", nName="${nName}", amount="${amount}"`);
                if (nName === "純損益に振り替えられることのない項目" || nName.includes("振り替えられることのない項目")) {
                    currentPLSubsection = "純損益に振り替えられることのない項目";
                    console.log(`[PL中間セクション] ${nName} -> ${currentPLSubsection}`);
                } else if (nName === "純損益に振り替えられる可能性のある項目" || nName.includes("振り替えられる可能性のある項目")) {
                    currentPLSubsection = "純損益に振り替えられる可能性のある項目";
                    console.log(`[PL中間セクション] ${nName} -> ${currentPLSubsection}`);
                } else if (nName.includes("当期利益の帰属") || nName.includes("親会社株主に係る当期利益")) {
                    currentPLSubsection = "当期利益の帰属";
                } else if (nName.includes("1株当たり") || nName.includes("１株当たり")) {
                    currentPLSubsection = "1株当たり当期利益";
                } else if (nName.includes("当期包括利益の帰属")) {
                    currentPLSubsection = "当期包括利益の帰属";
                } else if (nName === "その他の包括利益" || nName.includes("その他の包括利益合計")) {
                    currentPLSubsection = "その他の包括利益";
                } else if (nName.includes("当期包括利益") || nName.includes("税引後その他の包括利益")) {
                    // その他の包括利益セクションを抜けたか確認
                    if (!nName.includes("帰属") && nName !== "親会社株主に係る当期包括利益") {
                        currentPLSubsection = "";
                        console.log(`[PL中間セクション] セクション終了: ${nName}`);
                    }
                }
            }

            if (landmarkChanged && isSectionHeader) {
                // セクション見出しそのものの場合は、スタックを空にして本人のみ入れる
                hierarchyStack = [];
            } else if (nName.includes("合計") && !isSectionHeader && !isNonHierarchical) {
                // 「OO合計」などのデータ行は、多くの場合そのセクションの最後なので、スタックをポップする
                // ただし、階層化されていない期間はスタックをポップしない（ランドマークで管理するため）
                while (hierarchyStack.length > 0 && hierarchyStack[hierarchyStack.length - 1].level >= indentLevel) {
                    hierarchyStack.pop();
                }
            }

            // 階層スタックの管理: インデントレベルに基づいて親子関係を管理
            // ヘッダー、データ行どちらもスタックを更新
            while (hierarchyStack.length > 0 && hierarchyStack[hierarchyStack.length - 1].level >= indentLevel) {
                hierarchyStack.pop();
            }

            // ユニークキーの生成: ランドマークを最優先のアンカーにする
            const pathParts = [];

            // 階層化されていない期間は、シンプルにランドマーク+正規化名のみで統一
            const normalizedCol0 = normalizeKey(col0);

            if (isNonHierarchical) {
                // 階層化なしの場合: ランドマークのみ使用（スタックは使わない）
                if (currentLandmarks.major) pathParts.push(currentLandmarks.major);

                // 親セクションを名前から推測（財務諸表の種類に応じて異なるロジック）
                let parentSections = []; // 複数階層の親を格納

                if (currentBaseType === "連結貸借対照表") {
                    // 貸借対照表: 流動/非流動資産・負債の判定
                    if (normalizedCol0.includes("非流動資産")) {
                        parentSections.push("非流動資産");
                    } else if (normalizedCol0.includes("流動資産")) {
                        parentSections.push("流動資産");
                    } else if (normalizedCol0.includes("固定資産")) {
                        parentSections.push("固定資産");
                    } else if (normalizedCol0.includes("非流動負債")) {
                        parentSections.push("非流動負債");
                    } else if (normalizedCol0.includes("流動負債")) {
                        parentSections.push("流動負債");
                    } else if (normalizedCol0.includes("固定負債")) {
                        parentSections.push("固定負債");
                    } else if (currentLandmarks.sub) {
                        parentSections.push(currentLandmarks.sub);
                    }
                } else if (currentBaseType === "連結損益計算書") {
                    // 損益計算書: セクションごとに親を推測
                    if (normalizedCol0.includes("1株当たり") || normalizedCol0.includes("１株当たり")) {
                        parentSections.push("1株当たり当期利益");
                    } else if (normalizedCol0.includes("その他の包括利益")) {
                        parentSections.push("その他の包括利益");
                        // currentPLSubsectionを使用（事前に検出済み）
                        if ((currentPLSubsection === "純損益に振り替えられることのない項目" || currentPLSubsection === "純損益に振り替えられる可能性のある項目") && !normalizedCol0.includes("税引後その他の包括利益")) {
                            parentSections.push(currentPLSubsection);
                            console.log(`[パス構築] ${normalizedCol0} に中間セクション追加: ${currentPLSubsection}`);
                        }
                    } else if (normalizedCol0 === "当期利益" && currentPLSubsection === "当期利益の帰属") {
                        parentSections.push("当期利益の帰属");
                    } else if (normalizedCol0 === "当期包括利益" && currentPLSubsection === "当期包括利益の帰属") {
                        parentSections.push("当期包括利益の帰属");
                    } else if (normalizedCol0.includes("当期包括利益の帰属") ||
                        (normalizedCol0.includes("親会社") && normalizedCol0.includes("当期包括利益")) ||
                        (normalizedCol0.includes("非支配") && normalizedCol0.includes("当期包括利益"))) {
                        parentSections.push("当期包括利益の帰属");
                    } else if (normalizedCol0 === "当期包括利益") {
                        // OCI などのセクション外、ルートに戻す
                    } else if (currentPLSubsection === "純損益に振り替えられることのない項目" || currentPLSubsection === "純損益に振り替えられる可能性のある項目") {
                        parentSections.push("その他の包括利益");
                        parentSections.push(currentPLSubsection);
                    } else if (normalizedCol0.includes("当期利益の帰属") ||
                        (normalizedCol0.includes("親会社") && normalizedCol0.includes("所有者")) ||
                        (normalizedCol0.includes("非支配持分") && !normalizedCol0.includes("当期包括利益"))) {
                        parentSections.push("当期利益の帰属");
                    } else if (currentLandmarks.sub) {
                        parentSections.push(currentLandmarks.sub);
                    }
                } else {
                    // その他の財務諸表
                    if (currentLandmarks.sub) {
                        parentSections.push(currentLandmarks.sub);
                    }
                }

                // 親セクションをパスに追加 (重複排除)
                parentSections.forEach(ps => {
                    if (ps && !pathParts.includes(ps)) pathParts.push(ps);
                });

                // 本人の名前を追加（ランドマークや親セクションと同じ場合は除外）
                const isAlreadyInPath = pathParts.some(p => normalizeKey(p) === normalizedCol0);
                if (normalizedCol0 !== currentLandmarks.major && !isAlreadyInPath) {
                    pathParts.push(normalizedCol0);
                }
            } else {
                // 階層化ありの場合: 従来通りスタックを使用
                if (currentLandmarks.major) pathParts.push(currentLandmarks.major);
                if (currentLandmarks.sub) pathParts.push(currentLandmarks.sub);

                // スタック内のヘッダー名を追加 (ランドマークと重複しないように)
                hierarchyStack.forEach(h => {
                    const nh = normalizeKey(h.name);
                    if (nh !== currentLandmarks.major && nh !== currentLandmarks.sub) {
                        pathParts.push(nh);
                    }
                });

                // 本人の名前がまだパスに含まれていない場合のみ追加 (ヘッダー自身の場合など)
                if (pathParts.length === 0 || pathParts[pathParts.length - 1] !== normalizedCol0) {
                    pathParts.push(normalizedCol0);
                }
            }

            // 会計基準を決定 (後でまとめて年度ごとに上書き修正するが、一旦設定)
            let standard = yearStandards[currentYear] || "J-GAAP";

            // ユニークキーを「表名称|フルパス」にする（_headerサフィックスなし）
            let uniqueKey = `${currentBaseType}|${pathParts.join("|")}`;

            // デバッグ用
            if (isNonHierarchical && currentBaseType === "連結損益計算書" && (col0.includes("売上") || col0.includes("当期利益") || col0.includes("包括利益"))) {
                console.log(`[非階層化/PL] ${currentYear}: "${col0}" -> "${uniqueKey}"`);
            }
            if (!isNonHierarchical && currentBaseType === "連結損益計算書" && (col0.includes("売上") || col0.includes("当期利益") || col0.includes("包括利益"))) {
                console.log(`[階層化あり/PL] ${currentYear}: "${col0}" -> "${uniqueKey}"`);
            }

            // 同じ項目が既に存在するか確認
            const baseKeyExists = dictData[currentBaseType] && dictData[currentBaseType][uniqueKey];
            const headerKeyExists = dictData[currentBaseType] && dictData[currentBaseType][uniqueKey + "_header"];

            // 見出しとデータ行の統合ロジック
            if (isSectionHeader) {
                // 現在の行は見出し（金額なし）
                if (baseKeyExists) {
                    // 既にデータ行として登録済み → データ行を優先して見出しは無視
                    // ただし、初出が見出しの場合は統合
                    const existingIsHeader = dictIsHeader[currentBaseType] && dictIsHeader[currentBaseType][uniqueKey];
                    if (!existingIsHeader) {
                        // データ行が先に登録されているので、この見出しはスキップ
                        continue;
                    }
                    // 既存が見出しの場合は、その見出しが出現した年度を追加
                    if (currentYear) {
                        dictIsHeader[currentBaseType][uniqueKey].add(currentYear);
                    }
                } else if (headerKeyExists) {
                    // 既に見出しとして登録済み
                    uniqueKey = uniqueKey + "_header";
                    if (currentYear) {
                        dictIsHeader[currentBaseType][uniqueKey].add(currentYear);
                    }
                } else {
                    // 初出の見出し → とりあえず見出しとして登録（後でデータ行が来たら統合）
                    // 見出しサフィックスは付けない（データ行と統合するため）
                    if (!dictIsHeader[currentBaseType]) dictIsHeader[currentBaseType] = {};
                    dictIsHeader[currentBaseType][uniqueKey] = new Set();
                    if (currentYear) {
                        dictIsHeader[currentBaseType][uniqueKey].add(currentYear);
                    }
                }
            } else {
                // 現在の行はデータ行（金額あり）
                if (headerKeyExists) {
                    // 既に見出しとして登録済み → 見出しをデータ行に統合
                    // 見出しのデータを削除して、データ行として再登録
                    const headerKey = uniqueKey + "_header";
                    const headerIndex = dictItemsOrder[currentBaseType].indexOf(headerKey);
                    if (headerIndex !== -1) {
                        // 見出しを削除
                        dictItemsOrder[currentBaseType].splice(headerIndex, 1);
                        delete dictData[currentBaseType][headerKey];
                        delete dictIsHeader[currentBaseType][headerKey];
                        delete dictItemNames[currentBaseType][headerKey];
                        delete dictItemFirstSeen[currentBaseType][headerKey];
                    }
                    // データ行として登録（baseKeyで）
                }
                // baseKeyExistsの場合は、既存のデータに追加年度を加える
            }

            // 表示名は最初に出現したものを優先して保持
            if (!dictItemNames[currentBaseType][uniqueKey]) {
                dictItemNames[currentBaseType][uniqueKey] = itemName;
                if (!dictIsHeader[currentBaseType]) dictIsHeader[currentBaseType] = {};
                // 見出しフラグ（Set）の初期化
                if (isSectionHeader) {
                    dictIsHeader[currentBaseType][uniqueKey] = new Set();
                    if (currentYear) {
                        dictIsHeader[currentBaseType][uniqueKey].add(currentYear);
                    }
                }
            } else if (!isSectionHeader && dictIsHeader[currentBaseType][uniqueKey]) {
                // 既存が見出しだったが、今回データ行が来た → 見出しフラグを解除（Setを削除）
                delete dictIsHeader[currentBaseType][uniqueKey];
            }

            // 項目の順序を管理: 初出位置を記録し、適切な場所に挿入
            if (!dictItemsOrder[currentBaseType].includes(uniqueKey)) {
                // 初めて見る項目の場合、出現情報を記録
                dictItemFirstSeen[currentBaseType][uniqueKey] = {
                    yearIndex: yearIndex,
                    rowIndex: rowIdx
                };

                // 挿入位置を決定
                let insertIndex = dictItemsOrder[currentBaseType].length;

                // 親要素（スタックの最後の要素）を取得
                const parentKey = hierarchyStack.length > 0
                    ? hierarchyStack[hierarchyStack.length - 1].uniqueKey
                    : null;

                if (parentKey && dictItemsOrder[currentBaseType].includes(parentKey)) {
                    // 親要素が存在する場合、その子要素の最後に挿入
                    const parentIndex = dictItemsOrder[currentBaseType].indexOf(parentKey);

                    // 親の直後から、同じ親を持つ兄弟要素を探す
                    let lastSiblingIndex = parentIndex;
                    for (let i = parentIndex + 1; i < dictItemsOrder[currentBaseType].length; i++) {
                        const siblingKey = dictItemsOrder[currentBaseType][i];

                        // 同じ親の子要素かどうかをパスの深さで判断
                        const siblingPath = siblingKey.split('|');
                        const parentPath = parentKey.split('|');
                        const siblingHasParent = siblingPath.length > parentPath.length;

                        if (siblingHasParent) {
                            lastSiblingIndex = i;
                        } else {
                            // 同じレベルまたは上のレベルに達したら終了
                            break;
                        }
                    }

                    insertIndex = lastSiblingIndex + 1;
                } else {
                    // 親が見つからない場合、同じレベルの最後に追加
                    insertIndex = dictItemsOrder[currentBaseType].length;
                }

                dictItemsOrder[currentBaseType].splice(insertIndex, 0, uniqueKey);
            }

            // 現在の項目をスタックに追加（次の項目の親候補として）
            hierarchyStack.push({ level: indentLevel, name: col0, uniqueKey: uniqueKey });

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
    let foundIFRS = false;
    sortedYears.forEach(year => {
        if (yearStandards[year] === "IFRS") {
            foundIFRS = true;
        }
        if (foundIFRS) {
            yearStandards[year] = "IFRS";
        } else {
            if (!yearStandards[year]) yearStandards[year] = "J-GAAP";
        }
    });

    return {
        dictData,
        dictYears: sortedYears,
        dictItemsOrder,
        dictItemNames,
        yearStandards,
        dictIsHeader // 追加
    };
}

function generateExcelWorkbook(parsedData) {
    const wb = XLSX.utils.book_new();
    const sortedYears = parsedData.dictYears;
    const yearStandards = parsedData.yearStandards;
    const dictIsHeader = parsedData.dictIsHeader || {}; // 追加

    Object.keys(parsedData.dictData).forEach(baseType => {
        // 日本基準とIFRSでデータを分ける
        const standards = ["J-GAAP", "IFRS"];

        standards.forEach(std => {
            const targetYears = sortedYears.filter(y =>
                parsedData.yearStandards[y] === std
            );
            if (targetYears.length === 0) return;

            // シート名の決定
            let suffix = (std === "IFRS") ? "(IFRS)" : "(日本基準)";
            let sheetTitle = baseType + suffix;

            // 特殊対応: IFRSの貸借対照表は「連結財政状態計算書(IFRS)」とする
            if (baseType === "連結貸借対照表" && std === "IFRS") {
                sheetTitle = "連結財政状態計算書(IFRS)";
            }

            const items = parsedData.dictItemsOrder[baseType];

            // この基準（std）において、少なくとも1つの年度で値が存在する項目、またはその見出し項目を抽出
            let validItems = [];
            let currentHeader = null;
            let headerHasData = false;

            const sheetIsHeader = dictIsHeader[baseType] || {};

            items.forEach(uniqueKey => {
                const hasData = targetYears.some(year => parsedData.dictData[baseType][uniqueKey][year] !== undefined);
                const isHeader = sheetIsHeader[uniqueKey];

                if (isHeader) {
                    if (currentHeader && headerHasData) {
                        validItems.push(currentHeader);
                        // headerHasData 以下の項目を実際に追加するのは下のループ等ではなく即時。
                        // ただし、入れ子構造を完全に再現するのは難しいため、
                        // 「見出し」→「中身」→「次の見出し」の順で plain に追加していく。
                    }
                    currentHeader = uniqueKey;
                    headerHasData = false;
                } else if (hasData) {
                    if (currentHeader) {
                        validItems.push(currentHeader);
                        currentHeader = null;
                    }
                    validItems.push(uniqueKey);
                }
            });
            // 最後のヘッダー処理（もしデータがあったら）は不要。
            // 実際には items を2パスで見るのが確実。

            // 再考：シンプルに「データがある」or「直後にデータがある見出し」を判定
            // ただし、見出しは対象年度（targetYears）にデータがある場合のみ採用
            validItems = [];
            for (let i = 0; i < items.length; i++) {
                const key = items[i];
                const hasData = targetYears.some(year => parsedData.dictData[baseType][key][year] !== undefined);
                const isHeader = sheetIsHeader[key];

                if (hasData) {
                    validItems.push(key);
                } else if (isHeader) {
                    // 見出し項目が、対象年度（targetYears）のいずれかで「実際にその年度の見出し」として存在したかチェック
                    const headerYears = dictIsHeader[baseType][key];
                    const isValidHeaderInTargetYears = targetYears.some(year => headerYears.has(year));

                    if (isValidHeaderInTargetYears) {
                        // 次の見出しが出るまでに、対象年度（targetYears）にデータがある項目が1つでもあれば、この見出しを採用
                        let foundDataInTargetYears = false;
                        for (let j = i + 1; j < items.length; j++) {
                            const childKey = items[j];
                            const childIsHeader = sheetIsHeader[childKey];

                            // 次の同レベル以上の見出しが出たら終了
                            if (childIsHeader) {
                                // 階層の深さを判定（単純にキーの長さで判定）
                                const parentDepth = key.split('|').length;
                                const childDepth = childKey.split('|').length;
                                if (childDepth <= parentDepth) {
                                    break; // 同レベルまたは上のレベルの見出し → 終了
                                }
                            }

                            // この子項目が対象年度にデータを持つかチェック
                            if (targetYears.some(year => parsedData.dictData[baseType][childKey][year] !== undefined)) {
                                foundDataInTargetYears = true;
                                break;
                            }
                        }
                        if (foundDataInTargetYears) {
                            validItems.push(key);
                        }
                    }
                }
            }

            if (validItems.length === 0) return;

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
