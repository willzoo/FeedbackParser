function normalizeCellValue(cell) {
    if (cell === null || cell === undefined) {
        return "";
    }
    return String(cell).replace(/^\uFEFF/, "").trim();
}

function normalizeGrid(jsonData) {
    return jsonData.map((row) =>
        Array.isArray(row) ? row.map(normalizeCellValue) : []
    );
}

function isQualtricsThreeRowHeaderGrid(data) {
    if (!data || data.length < 4) {
        return false;
    }
    const marker = normalizeCellValue(data[2] && data[2][0]);
    return marker.includes("ImportId");
}

function getQuestionRowAndBodyRows(data) {
    const normalized = normalizeGrid(data);
    if (isQualtricsThreeRowHeaderGrid(normalized)) {
        return {
            questionRow: normalized[1],
            bodyRows: normalized.slice(3),
        };
    }
    return {
        questionRow: normalized[0],
        bodyRows: normalized.slice(1),
    };
}

// Excel column letter to index converter
function excelColumnToIndex(columnLetters) {
    let result = 0;
    for (let i = 0; i < columnLetters.length; i++) {
        const char = columnLetters[i].toUpperCase();
        result = result * 26 + (char.charCodeAt(0) - 'A'.charCodeAt(0) + 1);
    }
    return result - 1; // Convert to 0-based index
}

// Calculate Levenshtein distance between two strings
function levenshteinDistance(str1, str2) {
    const len1 = str1.length;
    const len2 = str2.length;
    
    // Create a matrix
    const matrix = [];
    for (let i = 0; i <= len1; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= len2; j++) {
        matrix[0][j] = j;
    }
    
    // Fill the matrix
    for (let i = 1; i <= len1; i++) {
        for (let j = 1; j <= len2; j++) {
            if (str1[i - 1] === str2[j - 1]) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j] + 1,     // deletion
                    matrix[i][j - 1] + 1,     // insertion
                    matrix[i - 1][j - 1] + 1 // substitution
                );
            }
        }
    }
    
    return matrix[len1][len2];
}

// Check if a name should be dropped (N/A or NA, case-insensitive)
function shouldDropName(name) {
    if (!name) return true;
    const upperName = name.toUpperCase().trim();
    return upperName === "N/A" || upperName === "NA";
}

// Lowercase normalization for name similarity (merging ignores capitalization)
function normalizeNameForComparison(name) {
    return String(name).toLowerCase();
}

function lowercaseLetterCount(s) {
    let n = 0;
    for (let i = 0; i < s.length; i++) {
        const c = s[i];
        if (c >= "a" && c <= "z") {
            n++;
        }
    }
    return n;
}

/** Prefer natural casing (e.g. "Caleigh" over "CALEIGH") when lengths match. */
function pickBetterDisplayName(a, b) {
    if (a === b) {
        return a;
    }
    const la = lowercaseLetterCount(a);
    const lb = lowercaseLetterCount(b);
    if (lb !== la) {
        return lb > la ? b : a;
    }
    if (b.length !== a.length) {
        return b.length > a.length ? b : a;
    }
    return a.localeCompare(b) <= 0 ? a : b;
}

// Calculate Levenshtein distance with special handling for different length names
function calculateNameDistance(name1, name2) {
    const n1 = normalizeNameForComparison(name1);
    const n2 = normalizeNameForComparison(name2);
    const len1 = n1.length;
    const len2 = n2.length;

    // If one name has X characters and another has X+Y characters,
    // calculate distance between first X characters of both
    if (len1 !== len2) {
        const minLen = Math.min(len1, len2);
        return levenshteinDistance(n1.substring(0, minLen), n2.substring(0, minLen));
    }

    return levenshteinDistance(n1, n2);
}

/**
 * Star-shaped clusters (same as legacy combine): seed order, each name joins only if distance < 3 from seed.
 * Returns only groups with length > 1.
 */
function findSimilarNameGroupsInQuestion(namesData) {
    if ("__no_name__" in namesData) {
        return [];
    }
    const names = Object.keys(namesData);
    if (names.length < 2) {
        return [];
    }
    const nameGroups = [];
    const processed = new Set();

    for (let i = 0; i < names.length; i++) {
        if (processed.has(names[i])) {
            continue;
        }
        const group = [names[i]];
        processed.add(names[i]);
        for (let j = i + 1; j < names.length; j++) {
            if (processed.has(names[j])) {
                continue;
            }
            const distance = calculateNameDistance(names[i], names[j]);
            if (distance < 3) {
                group.push(names[j]);
                processed.add(names[j]);
            }
        }
        if (group.length > 1) {
            nameGroups.push(group);
        }
    }
    return nameGroups;
}

/** Display label for a set of names (case-collapse + slash for distinct spellings). */
function formatMergedDisplayNameFromMembers(members) {
    const group = [...members].sort(
        (a, b) =>
            normalizeNameForComparison(a).localeCompare(normalizeNameForComparison(b)) ||
            a.localeCompare(b)
    );
    const byLower = new Map();
    for (const name of group) {
        const key = normalizeNameForComparison(name);
        const prev = byLower.get(key);
        if (!prev) {
            byLower.set(key, name);
        } else {
            byLower.set(key, pickBetterDisplayName(prev, name));
        }
    }
    return [...byLower.keys()]
        .sort((a, b) => a.localeCompare(b))
        .map((k) => byLower.get(k))
        .join("/");
}

function collectAllPersonNames(results) {
    const set = new Set();
    for (const question of Object.keys(results)) {
        const namesData = results[question];
        for (const name of Object.keys(namesData)) {
            if (name !== "__no_name__") {
                set.add(name);
            }
        }
    }
    return [...set].sort((a, b) => a.localeCompare(b));
}

class UnionFind {
    constructor(elements) {
        this.parent = new Map();
        for (const e of elements) {
            this.parent.set(e, e);
        }
    }
    find(x) {
        const p = this.parent.get(x);
        if (p === undefined) {
            return x;
        }
        if (p !== x) {
            this.parent.set(x, this.find(p));
        }
        return this.parent.get(x);
    }
    union(a, b) {
        if (!this.parent.has(a) || !this.parent.has(b)) {
            return;
        }
        const ra = this.find(a);
        const rb = this.find(b);
        if (ra !== rb) {
            this.parent.set(rb, ra);
        }
    }
    getComponents() {
        const rootToMembers = new Map();
        for (const x of this.parent.keys()) {
            const r = this.find(x);
            if (!rootToMembers.has(r)) {
                rootToMembers.set(r, []);
            }
            rootToMembers.get(r).push(x);
        }
        return [...rootToMembers.values()];
    }
}

/** Partition of all person names: either one chip each or union of per-question similarity groups. */
function buildInitialNameGroups(rawResults, useSimilarityAlgo) {
    const allNames = collectAllPersonNames(rawResults);
    if (allNames.length === 0) {
        return [];
    }
    if (!useSimilarityAlgo) {
        return allNames.map((n) => [n]);
    }
    const uf = new UnionFind(allNames);
    for (const question of Object.keys(rawResults)) {
        const groups = findSimilarNameGroupsInQuestion(rawResults[question]);
        for (const g of groups) {
            for (let k = 1; k < g.length; k++) {
                uf.union(g[0], g[k]);
            }
        }
    }
    return uf
        .getComponents()
        .map((g) => g.sort((a, b) => a.localeCompare(b)))
        .sort((a, b) => a[0].localeCompare(b[0]));
}

function deepCloneResults(obj) {
    return JSON.parse(JSON.stringify(obj));
}

function mergeQuestionBucketsByLabel(namesData, nameToLabel) {
    const newData = {};
    if (namesData.__no_name__) {
        newData.__no_name__ = { ...namesData.__no_name__ };
    }
    for (const [name, counts] of Object.entries(namesData)) {
        if (name === "__no_name__") {
            continue;
        }
        const label = nameToLabel.has(name) ? nameToLabel.get(name) : name;
        if (!newData[label]) {
            newData[label] = {};
        }
        for (const [k, v] of Object.entries(counts)) {
            newData[label][k] = (newData[label][k] || 0) + v;
        }
    }
    return newData;
}

/** Apply a global partition of names to all questions (merged display keys). */
function applyGroupsToResults(rawResults, groups) {
    const out = deepCloneResults(rawResults);
    const nameToLabel = new Map();
    for (const members of groups) {
        if (!members || members.length === 0) {
            continue;
        }
        const label = formatMergedDisplayNameFromMembers(members);
        for (const m of members) {
            nameToLabel.set(m, label);
        }
    }
    for (const question of Object.keys(out)) {
        out[question] = mergeQuestionBucketsByLabel(out[question], nameToLabel);
    }
    return out;
}

function wrapNameGroupsWithStableIds(groups2d) {
    return groups2d.map((members) => ({
        id: crypto.randomUUID(),
        members: [...members],
    }));
}

function nameGroupsToMemberArrays(groups) {
    return groups.map((g) => [...g.members]);
}

// Process Excel or Qualtrics CSV export data (raw person keys; grouping is applied separately)
function processExcelData(data) {
    const { questionRow, bodyRows } = getQuestionRowAndBodyRows(data);

    if (!questionRow || questionRow.length === 0) {
        throw new Error("No header row found. Check the file format.");
    }
    if (bodyRows.length === 0) {
        throw new Error("No data rows found after the header.");
    }

    // Column AR is where the pattern changes
    const AR_COLUMN_INDEX = excelColumnToIndex("AR"); // Should be 43 (0-based)
    
    // Special columns: BF-BH (scores only, no names) and BI (sentences only, no names)
    const BF_COLUMN_INDEX = excelColumnToIndex("BF"); // 57 (0-based)
    const BG_COLUMN_INDEX = excelColumnToIndex("BG"); // 58 (0-based)
    const BH_COLUMN_INDEX = excelColumnToIndex("BH"); // 59 (0-based)
    const BI_COLUMN_INDEX = excelColumnToIndex("BI"); // 60 (0-based)
    
    // 1. Identify Questions (legacy Excel: row 0; Qualtrics CSV: row 1 after short codes + ImportId row)
    const questionMap = {}; // Maps data_col_index -> {question: string, sectionType: string}
    const suffixToIgnore = " or NA";
    
    // Process columns before AR: Odd columns (1, 3, 5, ...) have questions
    for (let scoreCol = 1; scoreCol < Math.min(AR_COLUMN_INDEX, questionRow.length); scoreCol += 2) {
        const fullQuestion = normalizeCellValue(questionRow[scoreCol]);
        
        if (fullQuestion) {
            // RULE: Ignore the string " or NA" from the end if it appears
            let cleanedQuestion = fullQuestion;
            if (cleanedQuestion.endsWith(suffixToIgnore)) {
                cleanedQuestion = cleanedQuestion.slice(0, -suffixToIgnore.length);
            }
            
            questionMap[scoreCol] = { question: cleanedQuestion, sectionType: false }; // false = score section
        }
    }
    
    // Process columns from AR onwards: Even columns (44, 46, 48, ...) have questions
    // AR (column 43) is odd and contains names, so feedback/questions start at column 44
    // But skip BF-BH and BI which are handled separately
    const suffixToIgnoreAR = "or NA "; // For AR and after, remove "or NA " (with trailing space)
    for (let feedbackCol = AR_COLUMN_INDEX + 1; feedbackCol < questionRow.length; feedbackCol += 2) {
        // Skip BF-BH and BI columns (they're handled separately)
        if ([BF_COLUMN_INDEX, BG_COLUMN_INDEX, BH_COLUMN_INDEX, BI_COLUMN_INDEX].includes(feedbackCol)) {
            continue;
        }
        
        const fullQuestion = normalizeCellValue(questionRow[feedbackCol]);
        
        if (fullQuestion) {
            // RULE: For AR and after, remove "or NA " (with trailing space) from anywhere in the string
            let cleanedQuestion = fullQuestion;
            // Remove all occurrences of "or NA " (with trailing space)
            cleanedQuestion = cleanedQuestion.replace(/or NA /g, '').trim();
            // Also handle " or NA" without trailing space for backwards compatibility (only at the end)
            if (cleanedQuestion.endsWith(suffixToIgnore)) {
                cleanedQuestion = cleanedQuestion.slice(0, -suffixToIgnore.length).trim();
            }
            
            questionMap[feedbackCol] = { question: cleanedQuestion, sectionType: true }; // true = feedback section
        }
    }
    
    // Process BF-BH columns: Each column is one question, scores only (no names)
    // These are after AR, so use "or NA " removal rule
    for (const scoreOnlyCol of [BF_COLUMN_INDEX, BG_COLUMN_INDEX, BH_COLUMN_INDEX]) {
        if (scoreOnlyCol < questionRow.length) {
            const fullQuestion = normalizeCellValue(questionRow[scoreOnlyCol]);
            
            if (fullQuestion) {
                // RULE: For AR and after, remove "or NA " (with trailing space) from anywhere in the string
                let cleanedQuestion = fullQuestion;
                // Remove all occurrences of "or NA " (with trailing space)
                cleanedQuestion = cleanedQuestion.replace(/or NA /g, '').trim();
                // Also handle " or NA" without trailing space for backwards compatibility (only at the end)
                if (cleanedQuestion.endsWith(suffixToIgnore)) {
                    cleanedQuestion = cleanedQuestion.slice(0, -suffixToIgnore.length).trim();
                }
                
                questionMap[scoreOnlyCol] = { question: cleanedQuestion, sectionType: "score_only" };
            }
        }
    }
    
    // Process BI column: Sentences only (no names)
    // This is after AR, so use "or NA " removal rule
    if (BI_COLUMN_INDEX < questionRow.length) {
        const fullQuestion = normalizeCellValue(questionRow[BI_COLUMN_INDEX]);
        
        if (fullQuestion) {
            // RULE: For AR and after, remove "or NA " (with trailing space) from anywhere in the string
            let cleanedQuestion = fullQuestion;
            // Remove all occurrences of "or NA " (with trailing space)
            cleanedQuestion = cleanedQuestion.replace(/or NA /g, '').trim();
            // Also handle " or NA" without trailing space for backwards compatibility (only at the end)
            if (cleanedQuestion.endsWith(suffixToIgnore)) {
                cleanedQuestion = cleanedQuestion.slice(0, -suffixToIgnore.length).trim();
            }
            
            questionMap[BI_COLUMN_INDEX] = { question: cleanedQuestion, sectionType: "sentence_only" };
        }
    }
    
    if (Object.keys(questionMap).length === 0) {
        throw new Error("No questions found in the question header row. Check the file format.");
    }
    
    // Initialize the nested result dictionary
    // results[cleaned_question_key][name_key] = {score_or_feedback_string: count}
    // For no-name columns, use "__no_name__" as the name key
    const results = {};
    // Track which questions are from score sections (before AR) for tabular display
    const scoreSectionQuestions = new Set();
    const validScores = new Set(['C', 'D', 'E']);
    
    // 2. Process data rows (legacy Excel: from row 1; Qualtrics CSV: from row 3)
    for (let rowIndex = 0; rowIndex < bodyRows.length; rowIndex++) {
        const row = bodyRows[rowIndex];
        if (!row || row.length === 0 || !row.some((cell) => normalizeCellValue(cell) !== "")) {
            continue;
        }
        const maxCol = row.length - 1;
        
        // Process each question column
        for (const [dataColStr, questionData] of Object.entries(questionMap)) {
            const dataCol = parseInt(dataColStr);
            if (dataCol > maxCol) {
                continue;
            }
            
            const { question, sectionType } = questionData;
            
            // Initialize question in results if not exists
            if (!results[question]) {
                results[question] = {};
            }
            
            // Track score sections (before AR)
            if (sectionType === false) {
                scoreSectionQuestions.add(question);
            }
            
            if (sectionType === "score_only") {
                // BF-BH columns: Scores only, no names
                const score = normalizeCellValue(row[dataCol]);
                
                if (score) {
                    // Use special key "__no_name__" for no-name columns
                    if (!results[question]["__no_name__"]) {
                        results[question]["__no_name__"] = {};
                    }
                    results[question]["__no_name__"][score] = (results[question]["__no_name__"][score] || 0) + 1;
                }
            } else if (sectionType === "sentence_only") {
                // BI column: Sentences only, no names
                const sentence = normalizeCellValue(row[dataCol]);
                
                if (sentence) {
                    // Use special key "__no_name__" for no-name columns
                    if (!results[question]["__no_name__"]) {
                        results[question]["__no_name__"] = {};
                    }
                    results[question]["__no_name__"][sentence] = (results[question]["__no_name__"][sentence] || 0) + 1;
                }
            } else if (sectionType === true) {
                // From AR onwards: Even column = Feedback, Odd column = Name
                const nameCol = dataCol - 1; // Name is in the previous odd column
                const feedbackCol = dataCol; // Feedback is in this even column
                
                // Retrieve and clean data
                const name = nameCol >= 0 && nameCol < row.length ? normalizeCellValue(row[nameCol]) : "";
                const feedback =
                    feedbackCol < row.length ? normalizeCellValue(row[feedbackCol]) : "";
                
                // Validation and Aggregation - drop N/A or NA names
                if (name && feedback && !shouldDropName(name)) {
                    if (!results[question][name]) {
                        results[question][name] = {};
                    }
                    results[question][name][feedback] = (results[question][name][feedback] || 0) + 1;
                }
            } else {
                // Before AR: Odd column = Score, Even column = Name
                // Only count C, D, E scores (case-insensitive)
                const scoreCol = dataCol; // Score is in this odd column
                const nameCol = scoreCol + 1; // Name is in the next even column
                
                // Retrieve and clean data
                const name = nameCol < row.length ? normalizeCellValue(row[nameCol]) : "";
                const score = (scoreCol < row.length ? normalizeCellValue(row[scoreCol]) : "").toUpperCase();
                
                // Validation and Aggregation - only count C, D, E, and drop N/A or NA names
                if (name && score && validScores.has(score) && !shouldDropName(name)) {
                    if (!results[question][name]) {
                        results[question][name] = {};
                    }
                    results[question][name][score] = (results[question][name][score] || 0) + 1;
                }
            }
        }
    }
    
    return { results, scoreSectionQuestions };
}

// Display results
function displayResults(resultsData, scoreSectionQuestions) {
    const resultsDiv = document.getElementById("results");
    const resultsSection = document.getElementById("results-section");
    
    if (!resultsData || Object.keys(resultsData).length === 0) {
        resultsDiv.textContent = "No valid data found after processing the file.";
        resultsSection.style.display = "block";
        return;
    }
    
    let htmlOutput = "";
    let textOutput = "";
    
    // Sort questions for consistent output
    const sortedQuestions = Object.keys(resultsData).sort();
    
    for (const question of sortedQuestions) {
        const namesData = resultsData[question];
        const isScoreSection = scoreSectionQuestions.has(question);
        
        // Display the question header
        const questionHeader = `====================================================================================================\nQUESTION: ${question}\n====================================================================================================\n`;
        textOutput += questionHeader;
        htmlOutput += `<div class="question-header">QUESTION: ${question}</div>`;
        
        // Check if this is a no-name question (BF-BH or BI)
        if ("__no_name__" in namesData) {
            // This is a no-name question - display scores/sentences directly
            const scoresOrSentences = namesData["__no_name__"];
            
            // Sort scores/sentences alphabetically for consistent sub-display
            const sortedItems = Object.keys(scoresOrSentences).sort();
            
            for (const item of sortedItems) {
                const count = scoresOrSentences[item];
                const plural = count > 1 ? "s" : "";
                textOutput += `  -> '${item}' (${count} time${plural})\n`;
            }
            htmlOutput += `<div class="no-name-item">`;
            for (const item of sortedItems) {
                const count = scoresOrSentences[item];
                const plural = count > 1 ? "s" : "";
                htmlOutput += `<div>  -> '${item}' (${count} time${plural})</div>`;
            }
            htmlOutput += `</div>`;
        } else if (isScoreSection) {
            // Score section (before AR): Display in HTML table format with D, C, E order
            const sortedNames = Object.keys(namesData).sort();
            
            htmlOutput += `<table class="questions-table"><thead><tr><th>Name</th><th>D</th><th>C</th><th>E</th></tr></thead><tbody>`;
            
            for (const name of sortedNames) {
                const scores = namesData[name];
                const dCount = scores['D'] || 0;
                const cCount = scores['C'] || 0;
                const eCount = scores['E'] || 0;
                
                // HTML table row
                htmlOutput += `<tr><td>${name}</td><td>${dCount}</td><td>${cCount}</td><td>${eCount}</td></tr>`;
                
                // Plain text format for fallback
                textOutput += `  ${name.padEnd(5)} | D: ${String(dCount).padEnd(5)} | C: ${String(cCount).padEnd(5)} | E: ${String(eCount).padEnd(5)}\n`;
            }
            
            htmlOutput += `</tbody></table>`;
        } else {
            // Regular question with names (feedback sections from AR onwards)
            const sortedNames = Object.keys(namesData).sort();
            
            for (const name of sortedNames) {
                const scores = namesData[name];
                textOutput += `  PERSON: ${name}\n`;
                htmlOutput += `<div class="person-header">PERSON: ${name}</div>`;
                
                // Sort scores alphabetically for consistent sub-display
                const sortedScores = Object.keys(scores).sort();
                
                for (const score of sortedScores) {
                    // For AR sections, don't show count - just show the feedback text
                    textOutput += `    -> '${score}'\n`;
                    htmlOutput += `<div class="score-item">    -> '${score}'</div>`;
                }
                
                textOutput += `\n`; // Empty line after each person's details
                htmlOutput += `<br>`; // Empty line after each person's details
            }
        }
        
        textOutput += `\n`; // Double empty line after each question group
        htmlOutput += `<br><br>`; // Double empty line after each question group
    }
    
    // Use innerHTML for HTML content, but keep textContent as fallback
    resultsDiv.innerHTML = htmlOutput;
    resultsSection.style.display = "block";
}

/** @type {{ rawResults: object, scoreSectionQuestions: Set, groups: { id: string, members: string[] }[] } | null} */
let parserState = null;

function refreshResultsFromGroups() {
    if (!parserState) {
        return;
    }
    const merged = applyGroupsToResults(
        parserState.rawResults,
        nameGroupsToMemberArrays(parserState.groups)
    );
    displayResults(merged, parserState.scoreSectionQuestions);
}

function removeNameFromAllGroups(name) {
    for (const g of parserState.groups) {
        const i = g.members.indexOf(name);
        if (i >= 0) {
            g.members.splice(i, 1);
        }
    }
    parserState.groups = parserState.groups.filter((g) => g.members.length > 0);
}

function moveNameIntoGroupById(name, targetGroupId) {
    if (!parserState) {
        return;
    }
    removeNameFromAllGroups(name);
    const target = parserState.groups.find((g) => g.id === targetGroupId);
    if (target) {
        target.members.push(name);
        target.members.sort((a, b) => a.localeCompare(b));
    } else {
        parserState.groups.push({
            id: crypto.randomUUID(),
            members: [name],
        });
    }
}

function moveNameToNewSingleton(name) {
    if (!parserState) {
        return;
    }
    removeNameFromAllGroups(name);
    parserState.groups.push({
        id: crypto.randomUUID(),
        members: [name],
    });
}

function onNameChipDragStart(e) {
    const name = e.target.getAttribute("data-name");
    if (!name) {
        return;
    }
    e.dataTransfer.setData("application/x-feedback-parser-name", name);
    e.dataTransfer.effectAllowed = "move";
    e.target.classList.add("name-chip--dragging");
}

function onNameChipDragEnd(e) {
    e.target.classList.remove("name-chip--dragging");
}

function renderNameCombiner() {
    const section = document.getElementById("name-combiner-section");
    const workspace = document.getElementById("name-combiner-workspace");
    if (!section || !workspace) {
        return;
    }
    if (!parserState || collectAllPersonNames(parserState.rawResults).length === 0) {
        section.style.display = "none";
        workspace.innerHTML = "";
        return;
    }
    section.style.display = "block";
    workspace.innerHTML = "";

    parserState.groups.forEach((group) => {
        const box = document.createElement("div");
        box.className = "name-group";
        box.dataset.groupId = group.id;
        box.addEventListener("dragover", (ev) => {
            ev.preventDefault();
            ev.dataTransfer.dropEffect = "move";
        });
        box.addEventListener("drop", (ev) => {
            ev.preventDefault();
            const name = ev.dataTransfer.getData("application/x-feedback-parser-name");
            if (!name || !parserState) {
                return;
            }
            moveNameIntoGroupById(name, group.id);
            renderNameCombiner();
            refreshResultsFromGroups();
        });

        group.members.forEach((name) => {
            const chip = document.createElement("span");
            chip.className = "name-chip";
            chip.textContent = name;
            chip.setAttribute("data-name", name);
            chip.setAttribute("draggable", "true");
            chip.addEventListener("dragstart", onNameChipDragStart);
            chip.addEventListener("dragend", onNameChipDragEnd);
            box.appendChild(chip);
        });
        workspace.appendChild(box);
    });
}

function wireNameCombinerDropZone() {
    const ungroup = document.getElementById("name-combiner-ungroup");
    if (!ungroup || ungroup.dataset.wired === "1") {
        return;
    }
    ungroup.dataset.wired = "1";
    ungroup.addEventListener("dragover", (ev) => {
        ev.preventDefault();
        ev.dataTransfer.dropEffect = "move";
        ungroup.classList.add("name-combiner-ungroup--active");
    });
    ungroup.addEventListener("dragleave", () => {
        ungroup.classList.remove("name-combiner-ungroup--active");
    });
    ungroup.addEventListener("drop", (ev) => {
        ev.preventDefault();
        ungroup.classList.remove("name-combiner-ungroup--active");
        const name = ev.dataTransfer.getData("application/x-feedback-parser-name");
        if (!name || !parserState) {
            return;
        }
        moveNameToNewSingleton(name);
        renderNameCombiner();
        refreshResultsFromGroups();
    });
}

wireNameCombinerDropZone();

const combineNamesCheckboxEl = document.getElementById("combine-names-checkbox");
if (combineNamesCheckboxEl) {
    combineNamesCheckboxEl.addEventListener("change", function () {
        if (!parserState) {
            return;
        }
        const useSimilarity = combineNamesCheckboxEl.checked;
        const initialGroups = buildInitialNameGroups(parserState.rawResults, useSimilarity);
        parserState.groups = wrapNameGroupsWithStableIds(initialGroups);
        renderNameCombiner();
        refreshResultsFromGroups();
    });
}

// File input handler
document.getElementById("file-input").addEventListener("change", function(event) {
    const file = event.target.files[0];
    const statusDiv = document.getElementById("status");
    
    if (!file) {
        statusDiv.textContent = "File selection cancelled.";
        return;
    }
    
    statusDiv.textContent = `Processing: ${file.name}...`;
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            
            // Get the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to array of arrays
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                defval: "" 
            });
            
            // Convert all values to strings
            const dataAsStrings = jsonData.map(row => 
                row.map(cell => cell !== null && cell !== undefined ? String(cell) : "")
            );
            
            const combineNamesCheckbox = document.getElementById("combine-names-checkbox");
            const useSimilarity = combineNamesCheckbox ? combineNamesCheckbox.checked : true;

            const { results, scoreSectionQuestions } = processExcelData(dataAsStrings);
            const initialGroups = buildInitialNameGroups(results, useSimilarity);
            parserState = {
                rawResults: results,
                scoreSectionQuestions,
                groups: wrapNameGroupsWithStableIds(initialGroups),
            };

            renderNameCombiner();
            refreshResultsFromGroups();

            statusDiv.textContent = `Successfully processed: ${file.name}`;
        } catch (error) {
            parserState = null;
            const combinerSection = document.getElementById("name-combiner-section");
            if (combinerSection) {
                combinerSection.style.display = "none";
            }
            const combinerWorkspace = document.getElementById("name-combiner-workspace");
            if (combinerWorkspace) {
                combinerWorkspace.innerHTML = "";
            }
            statusDiv.textContent = `Error: ${error.message}`;
            console.error("Processing Error:", error);
            alert(`An error occurred during processing: ${error.message}`);
        }
    };
    
    reader.onerror = function() {
        statusDiv.textContent = "Error reading file.";
        alert("Error reading file.");
    };
    
    reader.readAsArrayBuffer(file);
});

