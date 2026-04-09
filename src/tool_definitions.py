"""Native function calling tool definitions for Gemini Office Agent.

Defines all Gemini function declarations that replace the manual JSON prompt
engineering approach. The model natively decides when to call tools vs respond
with plain text.
"""

from typing import List
from google.genai import types


# ── Operation enumerations per tool ─────────────────────────────────────────

_EXCEL_OPERATIONS = [
    "read", "create", "update", "add", "append",
    "delete_sheet", "delete_rows", "format", "auto_width",
    "formula", "merge", "add_chart", "update_chart", "move_chart",
    "resize_chart", "list_charts", "delete_chart", "sort",
    "remove_duplicates", "filter_and_copy", "insert_rows",
    "insert_columns", "freeze_panes", "unfreeze_panes", "find_and_replace",
]

_WORD_OPERATIONS = [
    "read", "create", "update", "add_heading", "add_table", "add_list",
    "add_paragraph", "update_paragraph", "format_paragraph", "delete_paragraph",
    "replace", "improve_text", "correct_grammar", "improve_clarity",
    "adjust_tone", "simplify_language", "rewrite_professional",
    "generate_summary", "extract_key_points", "create_resume",
    "generate_conclusions", "create_faq", "convert_list_to_table",
    "convert_table_to_list", "extract_tables_to_excel",
    "export_to_txt", "export_to_markdown", "export_to_html", "export_to_pdf",
    "analyze_word_count", "analyze_section_length", "get_document_statistics",
    "analyze_tone", "identify_jargon", "analyze_readability",
    "check_term_consistency", "analyze_document",
    "add_image", "add_image_at_position",
    "add_hyperlink", "add_hyperlink_to_paragraph",
    "add_header", "add_footer", "remove_header", "remove_footer",
    "set_page_margins", "set_page_size", "get_page_layout",
    "add_page_break", "add_section_break",
]

_PDF_OPERATIONS = [
    "read", "create", "split", "add_text", "get_info",
    "rotate", "extract_tables", "merge",
]

_POWERPOINT_OPERATIONS = [
    "read", "create", "update", "add", "delete_slide",
    "duplicate_slide", "add_textbox", "add_table", "replace",
]

# ── Parameter schema (flexible key-value pairs) ───────────────────────────

def _params_schema() -> types.Schema:
    """Generic parameters object schema for operation-specific arguments."""
    return types.Schema(
        type=types.Type.OBJECT,
        description=(
            "Operation-specific parameters as key-value pairs. "
            "Examples: {\"content\": \"text\"}, {\"old_text\": \"x\", \"new_text\": \"y\"}, "
            "{\"column\": \"A\", \"order\": \"asc\"}, {\"sheet\": \"Sheet1\"}, "
            "{\"elements\": [{\"type\": \"heading\", \"text\": \"Title\", \"level\": 1}]}"
        ),
        properties={
            "content":           types.Schema(type=types.Type.STRING),
            "old_text":          types.Schema(type=types.Type.STRING),
            "new_text":          types.Schema(type=types.Type.STRING),
            "sheet":             types.Schema(type=types.Type.STRING),
            "sheet_name":        types.Schema(type=types.Type.STRING),
            "column":            types.Schema(type=types.Type.STRING),
            "order":             types.Schema(type=types.Type.STRING),
            "text":              types.Schema(type=types.Type.STRING),
            "level":             types.Schema(type=types.Type.INTEGER),
            "index":             types.Schema(type=types.Type.INTEGER),
            "row":               types.Schema(type=types.Type.INTEGER),
            "rows":              types.Schema(type=types.Type.STRING),
            "headers":           types.Schema(type=types.Type.STRING),
            "items":             types.Schema(type=types.Type.STRING),
            "output_filename":   types.Schema(type=types.Type.STRING),
            "output_path":       types.Schema(type=types.Type.STRING),
            "tone":              types.Schema(type=types.Type.STRING),
            "title":             types.Schema(type=types.Type.STRING),
            "data":              types.Schema(type=types.Type.STRING),
            "filenames":         types.Schema(type=types.Type.STRING),
            "file_paths":        types.Schema(type=types.Type.STRING),
            "style":             types.Schema(type=types.Type.STRING),
            "chart_type":        types.Schema(type=types.Type.STRING),
            "chart_index":       types.Schema(type=types.Type.INTEGER),
            "position":          types.Schema(type=types.Type.STRING),
            "page":              types.Schema(type=types.Type.INTEGER),
            "start_page":        types.Schema(type=types.Type.INTEGER),
            "end_page":          types.Schema(type=types.Type.INTEGER),
            "pages":             types.Schema(type=types.Type.STRING),
            "page_size":         types.Schema(type=types.Type.STRING),
            "rotation":          types.Schema(type=types.Type.INTEGER),
            "scale":             types.Schema(type=types.Type.NUMBER),
            "width":             types.Schema(type=types.Type.NUMBER),
            "height":            types.Schema(type=types.Type.NUMBER),
            "x":                 types.Schema(type=types.Type.NUMBER),
            "y":                 types.Schema(type=types.Type.NUMBER),
            "opacity":           types.Schema(type=types.Type.NUMBER),
            "elements":          types.Schema(type=types.Type.STRING),
            "slides":            types.Schema(type=types.Type.STRING),
            "slide_index":       types.Schema(type=types.Type.INTEGER),
            "paragraph_index":   types.Schema(type=types.Type.INTEGER),
            "image_path":        types.Schema(type=types.Type.STRING),
            "url":               types.Schema(type=types.Type.STRING),
            "alignment":         types.Schema(type=types.Type.STRING),
            "bold":              types.Schema(type=types.Type.BOOLEAN),
            "italic":            types.Schema(type=types.Type.BOOLEAN),
            "font_size":         types.Schema(type=types.Type.NUMBER),
            "font_name":         types.Schema(type=types.Type.STRING),
            "start_row":         types.Schema(type=types.Type.INTEGER),
            "start_col":         types.Schema(type=types.Type.STRING),
            "count":             types.Schema(type=types.Type.INTEGER),
            "range":             types.Schema(type=types.Type.STRING),
            "cell_range":        types.Schema(type=types.Type.STRING),
            "operator":          types.Schema(type=types.Type.STRING),
            "value":             types.Schema(type=types.Type.STRING),
            "destination_sheet": types.Schema(type=types.Type.STRING),
            "destination_file":  types.Schema(type=types.Type.STRING),
            "has_header":        types.Schema(type=types.Type.BOOLEAN),
            "keep":              types.Schema(type=types.Type.STRING),
            "replace_existing":  types.Schema(type=types.Type.BOOLEAN),
            "bg_color":          types.Schema(type=types.Type.STRING),
            "font_color":        types.Schema(type=types.Type.STRING),
            "number_format":     types.Schema(type=types.Type.STRING),
            "wrap_text":         types.Schema(type=types.Type.BOOLEAN),
            "border":            types.Schema(type=types.Type.STRING),
            "match_case":        types.Schema(type=types.Type.BOOLEAN),
            "match_entire_cell": types.Schema(type=types.Type.BOOLEAN),
            "row_count":         types.Schema(
                type=types.Type.INTEGER,
                description="Number of rows to generate with sample data (used with create operation when user requests simulated/random data)",
            ),
            "columns":           types.Schema(
                type=types.Type.STRING,
                description="Comma-separated column names for data generation (e.g. 'Date,Revenue,Expenses,Profit')",
            ),
        },
    )


# ── Tool declarations ─────────────────────────────────────────────────────

def build_office_tools() -> List[types.Tool]:
    """Build the complete list of native tool definitions for office operations."""

    declarations = [

        # ── Utility ─────────────────────────────────────────────────────────
        types.FunctionDeclaration(
            name="list_files",
            description=(
                "List all Office files available in the current directory. "
                "Use this when the user asks what files exist, how many files there are, "
                "or which files are available."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                properties={},
            ),
        ),

        types.FunctionDeclaration(
            name="read_file",
            description=(
                "Read and return the full content of any Office file "
                "(Excel .xlsx, Word .docx, PDF .pdf, PowerPoint .pptx). "
                "Use this when the user wants to know what a file contains, "
                "asks you to read/analyse/summarise a file, or when you need "
                "the file content before modifying it."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "filename": types.Schema(
                        type=types.Type.STRING,
                        description="Exact file name including extension (e.g. 'report.pdf', 'data.xlsx')",
                    ),
                },
                required=["filename"],
            ),
        ),

        # ── Excel ─────────────────────────────────────────────────────────
        types.FunctionDeclaration(
            name="excel_operation",
            description=(
                "Execute any operation on an Excel (.xlsx) spreadsheet. "
                "Covers reading, creating, editing cells/rows/columns, sorting, "
                "charts, formulas, and formatting."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "filename": types.Schema(
                        type=types.Type.STRING,
                        description="Excel file name (e.g. 'data.xlsx')",
                    ),
                    "operation": types.Schema(
                        type=types.Type.STRING,
                        enum=_EXCEL_OPERATIONS,
                        description="Operation to perform on the Excel file",
                    ),
                    "parameters": _params_schema(),
                },
                required=["filename", "operation"],
            ),
        ),

        # ── Word ──────────────────────────────────────────────────────────
        types.FunctionDeclaration(
            name="word_operation",
            description=(
                "Execute any operation on a Word (.docx) document. "
                "Covers reading, creating, editing paragraphs/headings/tables, "
                "text improvement, summaries, exports, and analysis."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "filename": types.Schema(
                        type=types.Type.STRING,
                        description="Word file name (e.g. 'document.docx')",
                    ),
                    "operation": types.Schema(
                        type=types.Type.STRING,
                        enum=_WORD_OPERATIONS,
                        description="Operation to perform on the Word document",
                    ),
                    "parameters": _params_schema(),
                },
                required=["filename", "operation"],
            ),
        ),

        # ── PDF ───────────────────────────────────────────────────────────
        types.FunctionDeclaration(
            name="pdf_operation",
            description=(
                "Execute any operation on a PDF file. "
                "Covers reading, creating from scratch, splitting, merging, "
                "adding text/watermarks, rotating pages, and extracting tables."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "filename": types.Schema(
                        type=types.Type.STRING,
                        description="PDF file name (e.g. 'report.pdf')",
                    ),
                    "operation": types.Schema(
                        type=types.Type.STRING,
                        enum=_PDF_OPERATIONS,
                        description="Operation to perform on the PDF file",
                    ),
                    "parameters": _params_schema(),
                },
                required=["filename", "operation"],
            ),
        ),

        # ── PowerPoint ────────────────────────────────────────────────────
        types.FunctionDeclaration(
            name="powerpoint_operation",
            description=(
                "Execute any operation on a PowerPoint (.pptx) presentation. "
                "Covers reading, creating, editing slides, adding text boxes, "
                "duplicating or deleting slides."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "filename": types.Schema(
                        type=types.Type.STRING,
                        description="PowerPoint file name (e.g. 'slides.pptx')",
                    ),
                    "operation": types.Schema(
                        type=types.Type.STRING,
                        enum=_POWERPOINT_OPERATIONS,
                        description="Operation to perform on the PowerPoint file",
                    ),
                    "parameters": _params_schema(),
                },
                required=["filename", "operation"],
            ),
        ),
    ]

    return [types.Tool(function_declarations=declarations)]
