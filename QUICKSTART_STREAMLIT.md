# Quick Start Guide - Gemini Office Agent Streamlit Interface

## Prerequisites

1. Python 3.8 or higher installed
2. Gemini API key from Google AI Studio
3. Office files (.xlsx, .docx, .pptx) in a directory

## Setup (5 minutes)

### Step 1: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 2: Configure Environment

Create a `.env` file in the project root:

```env
GEMINI_API_KEY=your_api_key_here
ROOT_PATH=./
MODEL_NAME=gemini-2.0-flash-lite
```

Replace `your_api_key_here` with your actual Gemini API key.

### Step 3: Run the Application

```bash
streamlit run app.py
```

The app will open automatically in your browser at `http://localhost:8501`.

## First Use

### 1. Discover Files

When the app loads:
- It automatically scans for Office files in your `ROOT_PATH`
- Files appear in the sidebar under "📁 Arquivos Disponíveis"

### 2. Select Files (Optional)

In the sidebar:
- Check boxes next to files you want to work with
- Or click "✅ Selecionar Todos" to select all files
- Click "❌ Limpar Seleção" to deselect all

### 3. Enter a Command

In the main area:
- Type your request in natural language
- Example: "Create a new Excel file called sales.xlsx with columns for Date, Product, and Amount"

### 4. Submit and View Results

- Click "🚀 Enviar"
- Watch the progress indicators:
  - 📂 Discovering files
  - 📖 Reading content
  - 🤖 Calling Gemini API
  - ⚙️ Executing actions
- View the result (success ✅ or error ❌)
- Check modified files list

### 5. Review History

- Previous requests appear in the "📜 Histórico de Conversação" section
- Click on any entry to see details
- Click "🗑️ Limpar Histórico" to clear all history

## Example Commands

### Excel Operations
```
Create a new Excel file called budget.xlsx with sheets for Q1, Q2, Q3, Q4
Update cell A1 in sales.xlsx to "Total Revenue"
Add a new sheet called "Summary" to report.xlsx
```

### Word Operations
```
Create a new Word document called report.docx with a title "Annual Report"
Add a paragraph to meeting_notes.docx about the project status
Update the first paragraph in proposal.docx
```

### PowerPoint Operations
```
Create a PowerPoint presentation called pitch.pptx with 5 slides
Add a new slide to presentation.pptx with title "Conclusion"
Update slide 2 in quarterly_review.pptx
```

## Troubleshooting

### "No files discovered"
- Check that `ROOT_PATH` in `.env` points to the correct directory
- Ensure the directory contains `.xlsx`, `.docx`, or `.pptx` files
- Click "🔄 Atualizar Lista de Arquivos" to refresh

### "Error initializing Agent"
- Verify `GEMINI_API_KEY` is set correctly in `.env`
- Check that the API key is valid
- Ensure you have internet connectivity

### "API quota exceeded"
- You've reached your Gemini API usage limit
- Wait for the quota to reset or upgrade your plan

### App is slow
- Large files take longer to process
- Complex operations require more API calls
- Consider selecting only relevant files

## Tips for Best Results

1. **Be Specific**: Include file names and exact operations
   - Good: "Update cell B2 in sales.xlsx to 1500"
   - Bad: "Change a number"

2. **One Operation at a Time**: Break complex tasks into steps
   - Good: "Create budget.xlsx" then "Add Q1 sheet to budget.xlsx"
   - Bad: "Create budget.xlsx with 4 sheets and fill them with data"

3. **Use File Selection**: Select only relevant files to speed up processing
   - Reduces context size sent to Gemini
   - Faster response times

4. **Check History**: Review previous operations to avoid duplicates
   - History shows what was already done
   - Helps track changes

## Keyboard Shortcuts

- `Ctrl + Enter` (in text area): Submit request
- `Ctrl + R`: Refresh page
- `Esc`: Close modals/dialogs

## Next Steps

- Read the full [README_STREAMLIT.md](README_STREAMLIT.md) for detailed documentation
- Check the [Design Document](.kiro/specs/gemini-office-agent/design.md) for architecture details
- Review [Requirements](.kiro/specs/gemini-office-agent/requirements.md) for feature specifications

## Support

If you encounter issues:
1. Check the console output for error messages
2. Review the logs (if logging is enabled)
3. Verify your `.env` configuration
4. Ensure all dependencies are installed correctly

## Security Notes

- Never share your `GEMINI_API_KEY`
- The `.env` file is in `.gitignore` to prevent accidental commits
- The app only accesses files within `ROOT_PATH`
- All operations are validated before execution

---

**Ready to start?** Run `streamlit run app.py` and begin automating your Office tasks! 🚀
