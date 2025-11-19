# Feedback Parser

A web-based Excel feedback analyzer that processes structured feedback data from Excel files and displays organized results.

## Features

- 📊 Parse Excel files (.xlsx, .xls)
- 🔍 Analyze feedback data with complex column patterns
- 📈 Display organized results with counts
- 🌐 Works entirely in the browser (no server required)
- 📱 Responsive design for all devices

## How to Use

1. Open `index.html` in a web browser
2. Click "Choose Excel File" to select your Excel file
3. View the processed results below

## Excel File Format

The parser expects Excel files with the following structure:

- **Row 0**: Contains questions (one per column pair)
- **Row 1 onwards**: Data rows with names and scores/feedback
- **Column 0**: Ignored

### Column Patterns

**Before column AR:**
- Odd columns (1, 3, 5, ...): Scores
- Even columns (2, 4, 6, ...): Names

**From column AR onwards:**
- Even columns (AR, AT, AV, ...): Feedback sentences
- Odd columns (AS, AU, AW, ...): Names

**Special columns:**
- **BF-BH**: Scores only (no names), each column is one question
- **BI**: Sentences only (no names)

### Rules

- Full names are used for matching/display
- The string ' or NA' is automatically ignored from the end of questions if present
- Questions that are identical after removing ' or NA' are combined
- Results display counts of all unique scores/feedback given by each person for each question

## Hosting on GitHub Pages

1. Create a new repository on GitHub
2. Upload all files from this folder to the repository
3. Go to Settings → Pages
4. Select the branch (usually `main` or `master`)
5. Select the folder (usually `/ (root)`)
6. Click Save
7. Your site will be available at `https://yourusername.github.io/repository-name/`

## Files

- `index.html` - Main HTML file
- `style.css` - Styling
- `app.js` - JavaScript logic for parsing and display
- `README.md` - This file

## Technologies Used

- HTML5
- CSS3
- JavaScript (ES6+)
- [SheetJS (xlsx)](https://sheetjs.com/) - For reading Excel files

## License

Free to use and modify.

