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

// Calculate Levenshtein distance with special handling for different length names
function calculateNameDistance(name1, name2) {
    const len1 = name1.length;
    const len2 = name2.length;
    
    // If one name has X characters and another has X+Y characters,
    // calculate distance between first X characters of both
    if (len1 !== len2) {
        const minLen = Math.min(len1, len2);
        return levenshteinDistance(name1.substring(0, minLen), name2.substring(0, minLen));
    }
    
    return levenshteinDistance(name1, name2);
}

// Combine similar names in results
function combineSimilarNames(results) {
    // Process each question separately
    for (const question in results) {
        const namesData = results[question];
        
        // Skip if this is a no-name question
        if ("__no_name__" in namesData) {
            continue;
        }
        
        const names = Object.keys(namesData);
        if (names.length < 2) {
            continue; // Need at least 2 names to compare
        }
        
        // Find groups of similar names
        const nameGroups = [];
        const processed = new Set();
        
        for (let i = 0; i < names.length; i++) {
            if (processed.has(names[i])) {
                continue;
            }
            
            const group = [names[i]];
            processed.add(names[i]);
            
            // Find all names similar to this one
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
        
        // Combine names in each group
        for (const group of nameGroups) {
            // Sort group to ensure consistent ordering
            group.sort();
            
            // Create combined name: (First name)/(Next name)/(Next name)
            const combinedName = group.join('/');
            
            // Merge all data from group members into combined name
            const combinedData = {};
            for (const name of group) {
                const nameData = namesData[name];
                for (const key in nameData) {
                    if (combinedData[key]) {
                        combinedData[key] += nameData[key];
                    } else {
                        combinedData[key] = nameData[key];
                    }
                }
                // Remove original name
                delete namesData[name];
            }
            
            // Add combined name with merged data
            namesData[combinedName] = combinedData;
        }
    }
    
    return results;
}

// Process Excel data
function processExcelData(data) {
    if (data.length < 2) {
        throw new Error("Excel file must contain at least 2 rows (Row 0: Questions, Row 1+: Data).");
    }

    // Column AR is where the pattern changes
    const AR_COLUMN_INDEX = excelColumnToIndex("AR"); // Should be 43 (0-based)
    
    // Special columns: BF-BH (scores only, no names) and BI (sentences only, no names)
    const BF_COLUMN_INDEX = excelColumnToIndex("BF"); // 57 (0-based)
    const BG_COLUMN_INDEX = excelColumnToIndex("BG"); // 58 (0-based)
    const BH_COLUMN_INDEX = excelColumnToIndex("BH"); // 59 (0-based)
    const BI_COLUMN_INDEX = excelColumnToIndex("BI"); // 60 (0-based)
    
    // 1. Identify Questions (Row index 0)
    const questionRow = data[0];
    const questionMap = {}; // Maps data_col_index -> {question: string, sectionType: string}
    const suffixToIgnore = " or NA";
    
    // Process columns before AR: Odd columns (1, 3, 5, ...) have questions
    for (let scoreCol = 1; scoreCol < Math.min(AR_COLUMN_INDEX, questionRow.length); scoreCol += 2) {
        const fullQuestion = (questionRow[scoreCol] || "").toString().trim();
        
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
        
        const fullQuestion = (questionRow[feedbackCol] || "").toString().trim();
        
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
            const fullQuestion = (questionRow[scoreOnlyCol] || "").toString().trim();
            
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
        const fullQuestion = (questionRow[BI_COLUMN_INDEX] || "").toString().trim();
        
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
        throw new Error("No questions found in row 0. Check Excel format.");
    }
    
    // Initialize the nested result dictionary
    // results[cleaned_question_key][name_key] = {score_or_feedback_string: count}
    // For no-name columns, use "__no_name__" as the name key
    const results = {};
    // Track which questions are from score sections (before AR) for tabular display
    const scoreSectionQuestions = new Set();
    const validScores = new Set(['C', 'D', 'E']);
    
    // 2. Process Data Rows (starting from index 1)
    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
        const row = data[rowIndex];
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
                const score = (row[dataCol] || "").toString().trim();
                
                if (score) {
                    // Use special key "__no_name__" for no-name columns
                    if (!results[question]["__no_name__"]) {
                        results[question]["__no_name__"] = {};
                    }
                    results[question]["__no_name__"][score] = (results[question]["__no_name__"][score] || 0) + 1;
                }
            } else if (sectionType === "sentence_only") {
                // BI column: Sentences only, no names
                const sentence = (row[dataCol] || "").toString().trim();
                
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
                const name = (nameCol >= 0 && nameCol < row.length ? row[nameCol] : "").toString().trim();
                const feedback = (feedbackCol < row.length ? row[feedbackCol] : "").toString().trim();
                
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
                const name = (nameCol < row.length ? row[nameCol] : "").toString().trim();
                const score = (scoreCol < row.length ? row[scoreCol] : "").toString().trim().toUpperCase();
                
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
    
    // Combine similar names using Levenshtein distance
    combineSimilarNames(results);
    
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
            
            // Process the data
            const { results, scoreSectionQuestions } = processExcelData(dataAsStrings);
            
            // Display results
            displayResults(results, scoreSectionQuestions);
            
            statusDiv.textContent = `Successfully processed: ${file.name}`;
        } catch (error) {
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

