# Gemini Office Agent - Streamlit Interface

## Overview

This Streamlit application provides a web interface for the Gemini Office Agent, allowing users to manipulate Office files (Excel, Word, PowerPoint) through natural language commands.

## Features

### Task 14.1: Main Streamlit Application
- ✅ Entry point using `factory.create_agent()`
- ✅ Text input field for user prompts
- ✅ Submit button
- ✅ Results display area
- ✅ Error display with `st.error()`
- ✅ Spinner during processing
- ✅ Status messages for each workflow step:
  - 📂 Discovering files
  - 📖 Reading file content
  - 🤖 Calling Gemini API
  - ⚙️ Executing actions

### Task 14.2: File Selection Interface
- ✅ Display list of discovered Office files
- ✅ Checkboxes for file selection
- ✅ "Select All" button
- ✅ "Clear Selection" button
- ✅ Only selected files are passed to the Agent

### Task 14.3: Conversation History
- ✅ Session state for storing history
- ✅ Display previous prompts and results
- ✅ Clear history button
- ✅ Success/error indicators for each entry

## Installation

1. Ensure all dependencies are installed:
```bash
pip install -r requirements.txt
```

2. Configure your environment:
   - Copy `.env.example` to `.env`
   - Set your `GEMINI_API_KEY` in the `.env` file
   - Set `ROOT_PATH` to the directory containing your Office files

## Running the Application

Start the Streamlit app:

```bash
streamlit run app.py
```

The application will open in your default web browser at `http://localhost:8501`.

## Usage

### Basic Workflow

1. **File Discovery**: Click "🔄 Atualizar Lista de Arquivos" in the sidebar to scan for Office files
2. **File Selection** (Optional): Select specific files you want the Agent to work with
3. **Enter Command**: Type your natural language command in the text area
4. **Submit**: Click "🚀 Enviar" to process your request
5. **View Results**: See the operation result and any modified files
6. **Check History**: Review previous operations in the history panel

### Example Commands

- "Create a new Excel file with sales data"
- "Add a paragraph to the report.docx file"
- "Create a PowerPoint presentation with 3 slides about quarterly results"
- "Update cell A1 in budget.xlsx to 1000"

### Interface Layout

```
┌─────────────────────────────────────────────────────────┐
│  📄 Gemini Office Agent                                 │
├──────────────────┬──────────────────────────────────────┤
│  Sidebar         │  Main Content                        │
│  ⚙️ Configuration│  💬 Nova Solicitação                 │
│                  │  [Text Area for User Prompt]         │
│  📁 Files        │  [🚀 Enviar Button]                  │
│  □ file1.xlsx    │                                      │
│  □ file2.docx    │  📜 Histórico de Conversação         │
│  □ file3.pptx    │  [Previous requests and results]     │
│                  │                                      │
│  ✅ Select All   │                                      │
│  ❌ Clear        │                                      │
└──────────────────┴──────────────────────────────────────┘
```

## Architecture

The Streamlit app integrates with the existing Agent architecture:

```
Streamlit UI (app.py)
    ↓
Factory.create_agent()
    ↓
Agent.process_user_request()
    ↓
[FileScanner → GeminiClient → Tools]
    ↓
AgentResponse
    ↓
Display Results in UI
```

## Session State

The app maintains the following session state:
- `agent`: Initialized Agent instance
- `history`: List of previous requests and responses
- `discovered_files`: List of available Office files
- `selected_files`: List of user-selected files

## Error Handling

The app handles errors gracefully:
- Configuration errors (missing API key) are displayed on startup
- File discovery errors show informative messages
- Processing errors display detailed error information
- All errors are logged for debugging

## Progress Indicators

During request processing, the app shows:
1. 🔍 Processing request spinner
2. 📂 Discovering files status
3. 📖 Reading content status
4. 🤖 Calling Gemini API status
5. ⚙️ Executing actions status

## Requirements Validation

This implementation validates:
- **Requirements 1.1-1.5**: Complete Streamlit interface with input, progress, results, and error display
- **Design Document - Usability**: File selection interface for better user control
- **Requirements 1.3, 1.4**: Conversation history with session state

## Testing

Unit tests are provided in `tests/test_streamlit_app.py`. Note that testing Streamlit apps requires special mocking considerations.

To run tests:
```bash
pytest tests/test_streamlit_app.py -v
```

## Troubleshooting

### App won't start
- Check that `GEMINI_API_KEY` is set in your `.env` file
- Verify all dependencies are installed: `pip install -r requirements.txt`

### No files discovered
- Check that `ROOT_PATH` in `.env` points to a valid directory
- Ensure the directory contains `.xlsx`, `.docx`, or `.pptx` files
- Click "🔄 Atualizar Lista de Arquivos" to refresh

### API errors
- Verify your Gemini API key is valid
- Check your internet connection
- Review API quota limits

## Development

The app is structured as follows:

- `initialize_session_state()`: Sets up session variables
- `initialize_agent()`: Creates Agent using factory
- `discover_files()`: Scans for Office files
- `display_file_selector()`: Shows file selection UI
- `display_conversation_history()`: Shows previous interactions
- `process_user_request()`: Handles user commands
- `main()`: Entry point and layout

## Future Enhancements

Potential improvements:
- File upload capability
- Export conversation history
- Advanced filtering options
- Real-time file preview
- Multi-language support
