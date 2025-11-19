// Excel column letter to index converter
function excelColumnToIndex(columnLetters) {
    let result = 0;
    for (let i = 0; i < columnLetters.length; i++) {
        const char = columnLetters[i].toUpperCase();
        result = result * 26 + (char.charCodeAt(0) - 'A'.charCodeAt(0) + 1);
    }
    return result - 1; // Convert to 0-based index
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
    for (let feedbackCol = AR_COLUMN_INDEX + 1; feedbackCol < questionRow.length; feedbackCol += 2) {
        // Skip BF-BH and BI columns (they're handled separately)
        if ([BF_COLUMN_INDEX, BG_COLUMN_INDEX, BH_COLUMN_INDEX, BI_COLUMN_INDEX].includes(feedbackCol)) {
            continue;
        }
        
        const fullQuestion = (questionRow[feedbackCol] || "").toString().trim();
        
        if (fullQuestion) {
            // RULE: Ignore the string " or NA" from the end if it appears
            let cleanedQuestion = fullQuestion;
            if (cleanedQuestion.endsWith(suffixToIgnore)) {
                cleanedQuestion = cleanedQuestion.slice(0, -suffixToIgnore.length);
            }
            
            questionMap[feedbackCol] = { question: cleanedQuestion, sectionType: true }; // true = feedback section
        }
    }
    
    // Process BF-BH columns: Each column is one question, scores only (no names)
    for (const scoreOnlyCol of [BF_COLUMN_INDEX, BG_COLUMN_INDEX, BH_COLUMN_INDEX]) {
        if (scoreOnlyCol < questionRow.length) {
            const fullQuestion = (questionRow[scoreOnlyCol] || "").toString().trim();
            
            if (fullQuestion) {
                // RULE: Ignore the string " or NA" from the end if it appears
                let cleanedQuestion = fullQuestion;
                if (cleanedQuestion.endsWith(suffixToIgnore)) {
                    cleanedQuestion = cleanedQuestion.slice(0, -suffixToIgnore.length);
                }
                
                questionMap[scoreOnlyCol] = { question: cleanedQuestion, sectionType: "score_only" };
            }
        }
    }
    
    // Process BI column: Sentences only (no names)
    if (BI_COLUMN_INDEX < questionRow.length) {
        const fullQuestion = (questionRow[BI_COLUMN_INDEX] || "").toString().trim();
        
        if (fullQuestion) {
            // RULE: Ignore the string " or NA" from the end if it appears
            let cleanedQuestion = fullQuestion;
            if (cleanedQuestion.endsWith(suffixToIgnore)) {
                cleanedQuestion = cleanedQuestion.slice(0, -suffixToIgnore.length);
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
                
                // Validation and Aggregation
                if (name && feedback) {
                    if (!results[question][name]) {
                        results[question][name] = {};
                    }
                    results[question][name][feedback] = (results[question][name][feedback] || 0) + 1;
                }
            } else {
                // Before AR: Odd column = Score, Even column = Name
                const scoreCol = dataCol; // Score is in this odd column
                const nameCol = scoreCol + 1; // Name is in the next even column
                
                // Retrieve and clean data
                const name = (nameCol < row.length ? row[nameCol] : "").toString().trim();
                const score = (scoreCol < row.length ? row[scoreCol] : "").toString().trim();
                
                // Validation and Aggregation
                if (name && score) {
                    if (!results[question][name]) {
                        results[question][name] = {};
                    }
                    results[question][name][score] = (results[question][name][score] || 0) + 1;
                }
            }
        }
    }
    
    return results;
}

// Display results
function displayResults(resultsData) {
    const resultsDiv = document.getElementById("results");
    const resultsSection = document.getElementById("results-section");
    
    if (!resultsData || Object.keys(resultsData).length === 0) {
        resultsDiv.textContent = "No valid data found after processing the file.";
        resultsSection.style.display = "block";
        return;
    }
    
    let output = "";
    
    // Sort questions for consistent output
    const sortedQuestions = Object.keys(resultsData).sort();
    
    for (const question of sortedQuestions) {
        const namesData = resultsData[question];
        
        // Display the question
        output += `====================================================================================================\n`;
        output += `QUESTION: ${question}\n`;
        output += `====================================================================================================\n`;
        
        // Check if this is a no-name question (BF-BH or BI)
        if ("__no_name__" in namesData) {
            // This is a no-name question - display scores/sentences directly
            const scoresOrSentences = namesData["__no_name__"];
            
            // Sort scores/sentences alphabetically for consistent sub-display
            const sortedItems = Object.keys(scoresOrSentences).sort();
            
            for (const item of sortedItems) {
                const count = scoresOrSentences[item];
                const plural = count > 1 ? "s" : "";
                output += `  -> '${item}' (${count} time${plural})\n`;
            }
        } else {
            // Regular question with names
            const sortedNames = Object.keys(namesData).sort();
            
            for (const name of sortedNames) {
                const scores = namesData[name];
                output += `  PERSON: ${name}\n`;
                
                // Sort scores alphabetically for consistent sub-display
                const sortedScores = Object.keys(scores).sort();
                
                for (const score of sortedScores) {
                    const count = scores[score];
                    const plural = count > 1 ? "s" : "";
                    output += `    -> '${score}' (${count} time${plural})\n`;
                }
                
                output += `\n`; // Empty line after each person's details
            }
        }
        
        output += `\n`; // Double empty line after each question group
    }
    
    resultsDiv.textContent = output;
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
            const resultsData = processExcelData(dataAsStrings);
            
            // Display results
            displayResults(resultsData);
            
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

