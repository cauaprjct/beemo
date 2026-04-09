"""Response parser for Gemini Office Agent.

This module provides functionality to parse and validate responses from the Gemini API,
with support for JSON parsing and fallback to free-text parsing.
"""

import json
import re
from typing import Dict, Any, List, Optional

from src.exceptions import ValidationError
from src.logging_config import get_logger

logger = get_logger(__name__)


# Canonical operation names that Gemini may phrase differently
_OPERATION_ALIASES: Dict[str, str] = {
    # Chart operations
    'change_chart_type': 'update_chart',
    'change_chart':      'update_chart',
    'modify_chart':      'update_chart',
    'edit_chart':        'update_chart',
    'reposition_chart':  'move_chart',
    'relocate_chart':    'move_chart',
    'reposicionar_chart':'move_chart',
    'scale_chart':       'resize_chart',
    'enlarge_chart':     'resize_chart',
    'shrink_chart':      'resize_chart',
    # Generic operations
    'update_cell':       'update',
    'add_row':           'append',
    'add_slide':         'add',
    'update_slide':      'update',
    'extract_text':      'read',
    # Document creation (Gemini sometimes uses these variations)
    'create_document':   'create',
    'new_document':      'create',
    'make_document':     'create',
    'create_file':       'create',
    'new_file':          'create',
    'make_file':         'create',
    'generate_document': 'create',
    'generate_file':     'create',
    # Full-document content replacement (model sometimes uses these instead of create)
    'update_content':    'create',
    'replace_content':   'create',
    'set_content':       'create',
    'rewrite_document':  'create',
    'overwrite':         'create',
    'write_document':    'create',
    # Search/replace aliases
    'find_replace':      'replace',
    'search_replace':    'replace',
    'replace_text':      'replace',
    'substitute':        'replace',
    # Table update aliases (model uses these for Word table cell edits)
    'update_table':      'replace',
    'edit_table':        'replace',
    'modify_table':      'replace',
    'update_cell':       'replace',
    'edit_cell':         'replace',
    'update_row':        'replace',
    'edit_row':          'replace',
    # Format/style aliases
    'set_style':         'format_paragraph',
    'apply_style':       'format_paragraph',
    'change_style':      'format_paragraph',
    'change_format':     'format_paragraph',
    'apply_format':      'format_paragraph',
    'format_text':       'format_paragraph',
    # Paragraph delete aliases
    'remove_paragraph':  'delete_paragraph',
    'remove':            'delete_paragraph',
    # Heading aliases
    'add_section':       'add_heading',
    'insert_heading':    'add_heading',
    # Header/footer aliases
    'set_header':        'add_header',
    'set_footer':        'add_footer',
    'insert_header':     'add_header',
    'insert_footer':     'add_footer',
    # Row insert aliases (no direct Word op — map to replace so batch survives)
    'insert_row':        'replace',
    'add_row':           'replace',
    'append_row':        'replace',
    # Read aliases (model often uses read_file, open_file, etc.)
    'read_file':         'read',
    'open_file':         'read',
    'get_content':       'read',
    'get_text':          'read',
    'fetch_content':     'read',
    # PDF operation aliases
    'create_pdf':        'create',
    'new_pdf':           'create',
    'watermark':         'add_text',
    'add_watermark':     'add_text',
    'add_overlay':       'add_text',
    'text_overlay':      'add_text',
    'merge_pdfs':        'merge',
    'merge_pdf':         'merge',
    'combine_pdfs':      'merge',
    'split_pdf':         'split',
    'extract_pages':     'split',
    'pdf_info':          'get_info',
    'pdf_metadata':      'get_info',
    'rotate_pages':      'rotate',
    'rotate_pdf':        'rotate',
}

# Known parameter renames per canonical operation
_PARAMETER_ALIASES: Dict[str, Dict[str, str]] = {
    # Word replace / table-update parameter variants
    'replace': {
        'old_value':     'old_text',
        'new_value':     'new_text',
        'find':          'old_text',
        'search':        'old_text',
        'replacement':   'new_text',
        'replace_with':  'new_text',
        'value':         'new_text',
        'text':          'new_text',
        'new_content':   'new_text',
        'old_content':   'old_text',
    },
    'update_chart': {
        'new_type':   'chart_type',
        'type':       'chart_type',
        'sheet_name': 'sheet',
        'index':      'chart_index',
    },
    'resize_chart': {
        'factor':        'scale',
        'multiplier':    'scale',
        'ratio':         'scale',
        'size':          'scale',
        'w':             'width',
        'h':             'height',
        'sheet_name':    'sheet',
        'index':         'chart_index',
    },
    'move_chart': {
        'anchor':        'position',
        'cell':          'position',
        'location':      'position',
        'target':        'position',
        'new_position':  'position',
        'new_anchor':    'position',
        'target_cell':   'position',
        'destination':   'position',
        'new_cell':      'position',
        'to':            'position',
        'move_to':       'position',
        'sheet_name':    'sheet',
        'index':         'chart_index',
    },
    'add_chart': {
        'sheet_name': 'sheet',
    },
    'delete_chart': {
        'sheet_name': 'sheet',
        'chart_name': 'identifier',
        'title':      'identifier',
    },
}


class ResponseParser:
    """Parses and validates responses from the Gemini API.
    
    This class handles JSON parsing with fallback to free-text parsing,
    and validates the structure of parsed responses.
    """
    
    @staticmethod
    def parse_response(response_text: str) -> Dict[str, Any]:
        """Parse Gemini response and extract action structure.
        
        Attempts to parse as JSON first, falls back to free-text parsing if JSON fails.
        
        Args:
            response_text: Raw text response from Gemini API
        
        Returns:
            Dictionary containing parsed actions and explanation
            
        Raises:
            ValidationError: If response cannot be parsed or is invalid
            
        Validates: Requirements 7.5, 7.6
        """
        # Try JSON parsing first
        try:
            parsed = ResponseParser._parse_json(response_text)
            logger.info("Successfully parsed response as JSON")
            return parsed
        except (json.JSONDecodeError, ValueError) as e:
            logger.warning(f"JSON parsing failed: {e}. Attempting free-text parsing.")
            
            # Fallback to free-text parsing
            try:
                parsed = ResponseParser._parse_free_text(response_text)
                logger.info("Successfully parsed response as free text")
                return parsed
            except Exception as fallback_error:
                logger.error(f"Free-text parsing also failed: {fallback_error}")
                raise ValidationError(
                    f"Failed to parse response. JSON error: {e}. "
                    f"Free-text error: {fallback_error}"
                )
    
    @staticmethod
    def _parse_json(response_text: str) -> Dict[str, Any]:
        """Parse response as JSON.
        
        Args:
            response_text: Raw response text
        
        Returns:
            Parsed JSON dictionary
            
        Raises:
            json.JSONDecodeError: If text is not valid JSON
            ValidationError: If JSON structure is invalid
        """
        # Try to extract JSON from markdown code blocks if present
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', response_text, re.DOTALL)
        if json_match:
            json_text = json_match.group(1)
        else:
            # Try to find JSON object in the text
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            if json_match:
                json_text = json_match.group(0)
            else:
                json_text = response_text
        
        # Pre-clean: remove Python list comprehensions that LLMs sometimes generate inside JSON arrays
        # e.g. [{...} for r in range(2, 102)] → []
        json_text = re.sub(r'\{[^{}]*\}\s+for\s+\w+\s+in\s+[^\]]+', '', json_text)
        # Remove trailing commas before ] or } which may result from the above cleanup
        json_text = re.sub(r',\s*([\]\}])', r'\1', json_text)

        # Parse JSON — with multi-pass sanitization on errors
        try:
            parsed = json.loads(json_text)
        except json.JSONDecodeError as first_err:
            # Pass 2: sanitize common LLM JSON quirks and retry
            sanitized = ResponseParser._sanitize_json(json_text)
            try:
                parsed = json.loads(sanitized)
                logger.info("JSON reparado com sanitização (pass 2)")
            except json.JSONDecodeError:
                # Pass 3: truncation repair on sanitized text
                repaired = ResponseParser._try_repair_truncated_json(sanitized)
                if repaired is None:
                    # Last resort: try truncation repair on original text
                    repaired = ResponseParser._try_repair_truncated_json(json_text)
                if repaired is None:
                    raise first_err
                logger.warning(
                    f"JSON truncado detectado (char {first_err.pos}). "
                    f"Reparando estrutura automaticamente."
                )
                parsed = json.loads(repaired)
        
        # Normalize operation names and parameters before validation
        ResponseParser._normalize_parsed(parsed)
        
        # Validate structure
        ResponseParser._validate_structure(parsed)
        
        return parsed
    
    @staticmethod
    def _try_repair_truncated_json(json_text: str) -> Optional[str]:
        """Tenta reparar JSON truncado fechando estruturas abertas.

        Percorre o texto rastreando profundidade de chaves/colchetes e contexto
        de string para encontrar o último objeto de ação completo dentro do
        array 'actions', e reconstrói um envelope JSON válido em torno dele.

        Returns:
            String JSON reparada e válida, ou None se a reparação falhar.
        """
        in_string = False
        escape_next = False
        depth = 0
        last_complete_action_end = -1

        for i, ch in enumerate(json_text):
            if escape_next:
                escape_next = False
                continue
            if ch == '\\' and in_string:
                escape_next = True
                continue
            if ch == '"':
                in_string = not in_string
                continue
            if in_string:
                continue

            if ch in ('{', '['):
                depth += 1
            elif ch in ('}', ']'):
                prev_depth = depth
                depth -= 1
                # depth 3→2 with '}' means we just closed a top-level action object
                # (root={depth1}, actions=[{depth2}, action={depth3})
                if prev_depth == 3 and ch == '}':
                    last_complete_action_end = i

        if last_complete_action_end < 0:
            return None

        partial = json_text[:last_complete_action_end + 1]

        # Try each suffix to properly close the JSON envelope
        for suffix in (']}', '], "explanation": ""}'):
            candidate = partial + suffix
            try:
                json.loads(candidate)
                logger.info(
                    f"JSON reparado com sucesso usando suffix: '{suffix}'"
                )
                return candidate
            except json.JSONDecodeError:
                continue

        return None

    @staticmethod
    def _sanitize_json(json_text: str) -> str:
        """Sanitiza JSON com erros comuns gerados por LLMs.

        Corrige:
        - Booleans Python: True/False/None → true/false/null
        - Comentários JS: // linha e /* bloco */
        - Trailing commas antes de } ou ] (segunda passagem)
        - Aspas simples em chaves/valores simples
        """
        s = json_text

        # 1. Remove JS-style line comments (// ...)
        s = re.sub(r'//[^\n]*', '', s)

        # 2. Remove JS-style block comments (/* ... */)
        s = re.sub(r'/\*.*?\*/', '', s, flags=re.DOTALL)

        # 3. Fix Python booleans/None outside strings
        #    Uses word-boundary to avoid replacing inside quoted values
        s = re.sub(r'\bTrue\b', 'true', s)
        s = re.sub(r'\bFalse\b', 'false', s)
        s = re.sub(r'\bNone\b', 'null', s)

        # 4. Remove trailing commas before } or ] (aggressive re-run)
        for _ in range(5):
            new_s = re.sub(r',\s*([\]\}])', r'\1', s)
            if new_s == s:
                break
            s = new_s

        # 5. Replace single-quoted strings with double-quoted
        #    Only safe for simple key-value pairs; skip if apostrophes present
        def fix_single_quotes(m: re.Match) -> str:
            inner = m.group(1)
            if '"' in inner:
                return m.group(0)
            return f'"{inner}"'
        s = re.sub(r"'([^'\\]*)'", fix_single_quotes, s)

        return s

    # Default file extensions per tool
    _DEFAULT_EXTENSIONS: Dict[str, str] = {
        'word':        '.docx',
        'excel':       '.xlsx',
        'powerpoint':  '.pptx',
        'pdf':         '.pdf',
    }

    # Default base names per tool
    _DEFAULT_NAMES: Dict[str, str] = {
        'word':        'documento',
        'excel':       'planilha',
        'powerpoint':  'apresentacao',
        'pdf':         'documento',
    }

    @staticmethod
    def _infer_target_file(action: Dict[str, Any]) -> str:
        """Generate a sensible default target_file when the model omits it."""
        tool = action.get('tool', 'word')
        params = action.get('parameters', {}) or {}
        ext = ResponseParser._DEFAULT_EXTENSIONS.get(tool, '.docx')
        base = ResponseParser._DEFAULT_NAMES.get(tool, 'documento')

        # Use explicit hints from parameters
        for hint_key in ('file_name', 'filename', 'name', 'template_name', 'title'):
            hint = params.get(hint_key)
            if hint and isinstance(hint, str):
                # Strip any existing extension and re-attach the correct one
                hint = re.sub(r'\.[a-zA-Z]{2,5}$', '', hint.strip())
                if hint:
                    inferred = hint + ext
                    logger.info(f"Auto-inferred target_file from parameter '{hint_key}': {inferred}")
                    return inferred

        inferred = base + ext
        logger.info(f"Auto-inferred default target_file: {inferred}")
        return inferred

    @staticmethod
    def _normalize_parsed(parsed: Dict[str, Any]) -> None:
        """Normalize operation names and parameter keys in-place using alias tables."""
        for action in parsed.get('actions', []):
            if not isinstance(action, dict):
                continue
            # Normalize operation name
            op = action.get('operation', '')
            canonical = _OPERATION_ALIASES.get(op, op)
            if canonical != op:
                logger.info(f"Normalizing operation alias '{op}' → '{canonical}'")
                action['operation'] = canonical

            # Auto-fill missing or empty target_file
            if not action.get('target_file'):
                action['target_file'] = ResponseParser._infer_target_file(action)

            # Smart reroute: update_chart without chart_type but with size params → resize_chart
            if canonical == 'update_chart':
                params = action.get('parameters', {})
                size_keys = {'scale', 'width', 'height', 'factor', 'multiplier', 'ratio', 'size', 'w', 'h'}
                has_size = bool(size_keys & set(params.keys()))
                has_type = 'chart_type' in params or 'new_type' in params or 'type' in params
                if has_size and not has_type:
                    logger.info("Rerouting update_chart (no chart_type, has size params) → resize_chart")
                    action['operation'] = 'resize_chart'
                    canonical = 'resize_chart'

            # Normalize parameter keys
            param_map = _PARAMETER_ALIASES.get(canonical, {})
            if param_map and isinstance(action.get('parameters'), dict):
                params = action['parameters']
                for alias_key, canonical_key in param_map.items():
                    if alias_key in params and canonical_key not in params:
                        logger.info(f"Normalizing param '{alias_key}' → '{canonical_key}'")
                        params[canonical_key] = params.pop(alias_key)

    @staticmethod
    def _parse_free_text(response_text: str) -> Dict[str, Any]:
        """Parse response as free text when JSON parsing fails.
        
        Attempts to extract action information from natural language response.
        
        Args:
            response_text: Raw response text
        
        Returns:
            Dictionary with extracted action structure
            
        Raises:
            ValidationError: If essential information cannot be extracted
        """
        logger.info("Attempting to parse free-text response")
        
        actions = []
        explanation = response_text.strip()
        
        # Try to identify tool mentions
        tool_patterns = {
            'excel': r'\b(excel|spreadsheet|xlsx|workbook|sheet)\b',
            'word': r'\b(word|document|docx|paragraph)\b',
            'powerpoint': r'\b(powerpoint|presentation|pptx|slide)\b',
            'pdf': r'\b(pdf|portable document)\b'
        }
        
        detected_tool = None
        for tool, pattern in tool_patterns.items():
            if re.search(pattern, response_text, re.IGNORECASE):
                detected_tool = tool
                break
        
        # Try to identify operation
        operation_patterns = {
            'read': r'\b(read|open|view|show|display|get|extract)\b',
            'create': r'\b(create|new|make|generate)\b',
            'update': r'\b(update|modify|change|edit|set)\b',
            'add': r'\b(add|insert|append)\b'
        }
        
        detected_operation = None
        for operation, pattern in operation_patterns.items():
            if re.search(pattern, response_text, re.IGNORECASE):
                detected_operation = operation
                break
        
        # Try to extract file paths (handles Windows paths like C:\Users\... and quoted paths)
        file_path_pattern = r'"([^"]+\.(?:xlsx|docx|pptx|pdf))"'
        file_matches = re.findall(file_path_pattern, response_text, re.IGNORECASE)
        if not file_matches:
            # Fallback: unquoted paths, including Windows drive letter
            file_path_pattern = r'([a-zA-Z]:[\\\/][a-zA-Z0-9_/\\. -]+\.(?:xlsx|docx|pptx|pdf)|[a-zA-Z0-9_/\\.-]+\.(?:xlsx|docx|pptx|pdf))'
            file_matches = re.findall(file_path_pattern, response_text, re.IGNORECASE)
        target_file = file_matches[0] if file_matches else None
        if isinstance(target_file, tuple):
            target_file = target_file[0]
        
        # Build action if we have enough information
        if detected_tool and detected_operation:
            action = {
                'tool': detected_tool,
                'operation': detected_operation,
                'target_file': target_file or 'unknown',
                'parameters': {}
            }
            actions.append(action)
        else:
            # Cannot extract enough information
            raise ValidationError(
                "Could not extract sufficient action information from free-text response. "
                f"Detected tool: {detected_tool}, operation: {detected_operation}"
            )
        
        return {
            'actions': actions,
            'explanation': explanation
        }
    
    @staticmethod
    def _validate_structure(parsed: Dict[str, Any]) -> None:
        """Validate the structure of a parsed response.
        
        Args:
            parsed: Parsed response dictionary
        
        Raises:
            ValidationError: If structure is invalid
            
        Validates: Requirements 7.5
        """
        # Check required top-level fields
        if 'actions' not in parsed:
            raise ValidationError("Response missing required field: 'actions'")
        
        if not isinstance(parsed['actions'], list):
            raise ValidationError("Field 'actions' must be a list")
        
        if len(parsed['actions']) == 0:
            raise ValidationError("Field 'actions' cannot be empty")
        
        # Validate each action — skip invalid ones with a warning instead of
        # failing the entire batch (one bad action should not discard 9 good ones)
        valid_actions = []
        for i, action in enumerate(parsed['actions']):
            try:
                ResponseParser._validate_action(action, i)
                valid_actions.append(action)
            except ValidationError as ve:
                logger.warning(
                    f"Ação {i} ignorada por erro de validação (batch continua): {ve}"
                )

        if not valid_actions:
            raise ValidationError(
                "Nenhuma ação válida encontrada na resposta após filtragem."
            )

        parsed['actions'] = valid_actions
        
        # Explanation is optional but should be a string if present
        if 'explanation' in parsed and not isinstance(parsed['explanation'], str):
            raise ValidationError("Field 'explanation' must be a string")
    
    @staticmethod
    def _validate_action(action: Dict[str, Any], index: int) -> None:
        """Validate a single action structure.
        
        Args:
            action: Action dictionary to validate
            index: Index of action in actions list (for error messages)
        
        Raises:
            ValidationError: If action structure is invalid
            
        Validates: Requirements 7.5, 7.6
        """
        if not isinstance(action, dict):
            raise ValidationError(f"Action at index {index} must be a dictionary")
        
        # Check required fields
        required_fields = ['tool', 'operation', 'target_file', 'parameters']
        for field in required_fields:
            if field not in action:
                raise ValidationError(
                    f"Action at index {index} missing required field: '{field}'"
                )
        
        # Validate tool
        valid_tools = ['excel', 'word', 'powerpoint', 'pdf']
        if action['tool'] not in valid_tools:
            raise ValidationError(
                f"Action at index {index} has invalid tool: '{action['tool']}'. "
                f"Must be one of: {valid_tools}"
            )
        
        # Validate operation
        valid_operations = ['read', 'create', 'update', 'add', 'append', 'delete_sheet', 'delete_rows', 'format', 'auto_width', 'formula', 'merge', 'add_chart', 'update_chart', 'move_chart', 'resize_chart', 'list_charts', 'delete_chart', 'sort', 'remove_duplicates', 'filter_and_copy', 'insert_rows', 'insert_columns', 'freeze_panes', 'unfreeze_panes', 'find_and_replace', 'add_heading', 'add_table', 'add_list', 'add_paragraph', 'update_paragraph', 'format_paragraph', 'delete_paragraph', 'replace', 'improve_text', 'correct_grammar', 'improve_clarity', 'adjust_tone', 'simplify_language', 'rewrite_professional', 'generate_summary', 'extract_key_points', 'create_resume', 'generate_conclusions', 'create_faq', 'convert_list_to_table', 'convert_table_to_list', 'extract_tables_to_excel', 'export_to_txt', 'export_to_markdown', 'export_to_html', 'export_to_pdf', 'analyze_word_count', 'analyze_section_length', 'get_document_statistics', 'analyze_tone', 'identify_jargon', 'analyze_readability', 'check_term_consistency', 'analyze_document', 'delete_slide', 'duplicate_slide', 'add_textbox', 'split', 'add_text', 'get_info', 'rotate', 'extract_tables']
        if action['operation'] not in valid_operations:
            raise ValidationError(
                f"Action at index {index} has invalid operation: '{action['operation']}'. "
                f"Must be one of: {valid_operations}"
            )
        
        # Validate target_file
        if not isinstance(action['target_file'], str) or not action['target_file']:
            raise ValidationError(
                f"Action at index {index} has invalid target_file: must be a non-empty string"
            )
        
        # Validate parameters
        if not isinstance(action['parameters'], dict):
            raise ValidationError(
                f"Action at index {index} has invalid parameters: must be a dictionary"
            )
        
        # Validate operation-specific parameters
        ResponseParser._validate_operation_parameters(action, index)
    
    @staticmethod
    def _validate_operation_parameters(action: Dict[str, Any], index: int) -> None:
        """Validate operation-specific parameters.
        
        Args:
            action: Action dictionary to validate
            index: Index of action in actions list (for error messages)
        
        Raises:
            ValidationError: If operation-specific parameters are invalid
        """
        tool = action['tool']
        operation = action['operation']
        parameters = action['parameters']
        
        # Excel-specific validations
        if tool == 'excel':
            if operation == 'update':
                # Check if using update_range (multiple cells)
                if 'updates' in parameters:
                    if not isinstance(parameters['updates'], list):
                        raise ValidationError(
                            f"Action {index}: 'updates' must be a list"
                        )
                    if len(parameters['updates']) == 0:
                        raise ValidationError(
                            f"Action {index}: 'updates' list cannot be empty"
                        )
                    # Validate each update entry
                    for i, update in enumerate(parameters['updates']):
                        if not isinstance(update, dict):
                            raise ValidationError(
                                f"Action {index}: update at position {i} must be a dictionary"
                            )
                        required_update_fields = ['row', 'col', 'value']
                        for field in required_update_fields:
                            if field not in update:
                                raise ValidationError(
                                    f"Action {index}: update at position {i} missing required field '{field}'. "
                                    f"Expected format: {{'row': int, 'col': int, 'value': any}}"
                                )
                        # Validate types
                        if not isinstance(update['row'], int) or update['row'] < 1:
                            raise ValidationError(
                                f"Action {index}: update at position {i} has invalid 'row' (must be integer >= 1)"
                            )
                        if not isinstance(update['col'], int) or update['col'] < 1:
                            raise ValidationError(
                                f"Action {index}: update at position {i} has invalid 'col' (must be integer >= 1)"
                            )
                # Check if using single cell update
                elif 'row' in parameters or 'col' in parameters:
                    required_single_fields = ['sheet', 'row', 'col', 'value']
                    for field in required_single_fields:
                        if field not in parameters:
                            raise ValidationError(
                                f"Action {index}: single cell update missing required field '{field}'"
                            )
                else:
                    raise ValidationError(
                        f"Action {index}: update operation requires either 'updates' list or 'row'/'col'/'value' fields"
                    )
            
            elif operation == 'formula':
                # Check if using multiple formulas
                if 'formulas' in parameters:
                    if not isinstance(parameters['formulas'], list):
                        raise ValidationError(
                            f"Action {index}: 'formulas' must be a list"
                        )
                    # Validate each formula entry
                    for i, formula in enumerate(parameters['formulas']):
                        if not isinstance(formula, dict):
                            raise ValidationError(
                                f"Action {index}: formula at position {i} must be a dictionary"
                            )
                        required_formula_fields = ['row', 'col', 'formula']
                        for field in required_formula_fields:
                            if field not in formula:
                                raise ValidationError(
                                    f"Action {index}: formula at position {i} missing required field '{field}'"
                                )
                # Check if using single formula
                elif 'row' in parameters or 'col' in parameters:
                    required_single_fields = ['sheet', 'row', 'col', 'formula']
                    for field in required_single_fields:
                        if field not in parameters:
                            raise ValidationError(
                                f"Action {index}: single formula missing required field '{field}'"
                            )
            
            elif operation == 'format':
                if 'formatting' not in parameters:
                    raise ValidationError(
                        f"Action {index}: format operation requires 'formatting' parameter"
                    )
                formatting = parameters['formatting']
                if not isinstance(formatting, dict):
                    raise ValidationError(
                        f"Action {index}: 'formatting' must be a dictionary"
                    )
                if 'range' not in formatting:
                    raise ValidationError(
                        f"Action {index}: 'formatting' must include 'range' field (e.g., 'A1:C10')"
                    )
            
            elif operation == 'append':
                if 'rows' not in parameters:
                    raise ValidationError(
                        f"Action {index}: append operation requires 'rows' parameter"
                    )
                if not isinstance(parameters['rows'], list):
                    raise ValidationError(
                        f"Action {index}: 'rows' must be a list"
                    )
            
            elif operation == 'add_chart':
                if 'chart_config' not in parameters:
                    raise ValidationError(
                        f"Action {index}: add_chart operation requires 'chart_config' parameter"
                    )
                chart_config = parameters['chart_config']
                if not isinstance(chart_config, dict):
                    raise ValidationError(
                        f"Action {index}: 'chart_config' must be a dictionary"
                    )
                # Validate chart type
                if 'type' not in chart_config:
                    raise ValidationError(
                        f"Action {index}: 'chart_config' must include 'type' field"
                    )
                valid_chart_types = ['bar', 'column', 'line', 'pie', 'area', 'scatter']
                if chart_config['type'].lower() not in valid_chart_types:
                    raise ValidationError(
                        f"Action {index}: invalid chart type '{chart_config['type']}'. "
                        f"Must be one of: {', '.join(valid_chart_types)}"
                    )
                # Validate data specification
                has_data_range = 'data_range' in chart_config
                has_values = 'values' in chart_config
                if not has_data_range and not has_values:
                    raise ValidationError(
                        f"Action {index}: 'chart_config' must include either 'data_range' or 'values'"
                    )
            
            elif operation == 'update_chart':
                if 'chart_type' not in parameters:
                    raise ValidationError(
                        f"Action {index}: update_chart operation requires 'chart_type' parameter"
                    )
                valid_chart_types = ['bar', 'column', 'line', 'pie', 'area', 'scatter']
                if parameters['chart_type'].lower() not in valid_chart_types:
                    raise ValidationError(
                        f"Action {index}: invalid chart_type '{parameters['chart_type']}'. "
                        f"Must be one of: {', '.join(valid_chart_types)}"
                    )

            elif operation == 'move_chart':
                pass  # position validated at execution time for flexibility

            elif operation == 'resize_chart':
                pass  # scale/width/height validated at execution time

            elif operation == 'delete_chart':
                if 'identifier' not in parameters:
                    raise ValidationError(
                        f"Action {index}: delete_chart operation requires 'identifier' parameter"
                    )
                identifier = parameters['identifier']
                # Identifier must be int (index) or str (title)
                if not isinstance(identifier, (int, str)):
                    raise ValidationError(
                        f"Action {index}: 'identifier' must be int (index) or str (title), "
                        f"got {type(identifier).__name__}"
                    )
                if isinstance(identifier, str) and not identifier.strip():
                    raise ValidationError(
                        f"Action {index}: 'identifier' cannot be empty string"
                    )
            
            elif operation == 'sort':
                if 'sort_config' not in parameters:
                    raise ValidationError(
                        f"Action {index}: sort operation requires 'sort_config' parameter"
                    )
                sort_config = parameters['sort_config']
                if not isinstance(sort_config, dict):
                    raise ValidationError(
                        f"Action {index}: 'sort_config' must be a dictionary"
                    )
                # Validate column
                if 'column' not in sort_config:
                    raise ValidationError(
                        f"Action {index}: 'sort_config' must include 'column' field"
                    )
                # Validate order if present
                if 'order' in sort_config:
                    order = sort_config['order'].lower()
                    if order not in ['asc', 'desc']:
                        raise ValidationError(
                            f"Action {index}: invalid order '{sort_config['order']}'. "
                            f"Must be 'asc' or 'desc'"
                        )
            
            elif operation == 'remove_duplicates':
                if 'config' not in parameters:
                    raise ValidationError(
                        f"Action {index}: remove_duplicates operation requires 'config' parameter"
                    )
                config = parameters['config']
                if not isinstance(config, dict):
                    raise ValidationError(
                        f"Action {index}: 'config' must be a dictionary"
                    )
                # Validate keep if present
                if 'keep' in config:
                    keep = config['keep'].lower()
                    if keep not in ['first', 'last']:
                        raise ValidationError(
                            f"Action {index}: invalid keep '{config['keep']}'. "
                            f"Must be 'first' or 'last'"
                        )
            
            elif operation == 'filter_and_copy':
                if 'config' not in parameters:
                    raise ValidationError(
                        f"Action {index}: filter_and_copy operation requires 'config' parameter"
                    )
                config = parameters['config']
                if not isinstance(config, dict):
                    raise ValidationError(
                        f"Action {index}: 'config' must be a dictionary"
                    )
                # Validate required fields
                required_fields = ['column', 'operator', 'value']
                for field in required_fields:
                    if field not in config:
                        raise ValidationError(
                            f"Action {index}: 'config' must include '{field}' field"
                        )
                # Validate operator
                valid_operators = ['>', '<', '>=', '<=', '==', '!=', 'contains', 'starts_with', 'ends_with']
                if config['operator'] not in valid_operators:
                    raise ValidationError(
                        f"Action {index}: invalid operator '{config['operator']}'. "
                        f"Must be one of: {', '.join(valid_operators)}"
                    )
                # Validate destination
                if 'destination_sheet' not in config and 'destination_file' not in config:
                    raise ValidationError(
                        f"Action {index}: 'config' must include either 'destination_sheet' or 'destination_file'"
                    )
            
            elif operation == 'insert_rows':
                if 'start_row' not in parameters:
                    raise ValidationError(
                        f"Action {index}: insert_rows operation requires 'start_row' parameter"
                    )
                if not isinstance(parameters['start_row'], int) or parameters['start_row'] < 1:
                    raise ValidationError(
                        f"Action {index}: 'start_row' must be an integer >= 1"
                    )
                if 'count' in parameters:
                    if not isinstance(parameters['count'], int) or parameters['count'] < 1:
                        raise ValidationError(
                            f"Action {index}: 'count' must be an integer >= 1"
                        )
            
            elif operation == 'insert_columns':
                if 'start_col' not in parameters:
                    raise ValidationError(
                        f"Action {index}: insert_columns operation requires 'start_col' parameter"
                    )
                # start_col can be int or str (letter)
                if not isinstance(parameters['start_col'], (int, str)):
                    raise ValidationError(
                        f"Action {index}: 'start_col' must be an integer or string (column letter)"
                    )
                if isinstance(parameters['start_col'], int) and parameters['start_col'] < 1:
                    raise ValidationError(
                        f"Action {index}: 'start_col' must be >= 1 when using integer"
                    )
                if 'count' in parameters:
                    if not isinstance(parameters['count'], int) or parameters['count'] < 1:
                        raise ValidationError(
                            f"Action {index}: 'count' must be an integer >= 1"
                        )
            
            elif operation == 'freeze_panes':
                # Validate at least one parameter is provided
                if 'row' not in parameters and 'col' not in parameters:
                    raise ValidationError(
                        f"Action {index}: freeze_panes operation requires at least 'row' or 'col' parameter"
                    )
                # Validate row if present
                if 'row' in parameters:
                    if not isinstance(parameters['row'], int) or parameters['row'] < 1:
                        raise ValidationError(
                            f"Action {index}: 'row' must be an integer >= 1"
                        )
                # Validate col if present (can be int or str)
                if 'col' in parameters:
                    if not isinstance(parameters['col'], (int, str)):
                        raise ValidationError(
                            f"Action {index}: 'col' must be an integer or string (column letter)"
                        )
                    if isinstance(parameters['col'], int) and parameters['col'] < 1:
                        raise ValidationError(
                            f"Action {index}: 'col' must be >= 1 when using integer"
                        )
            
            elif operation == 'find_and_replace':
                # Validate required parameters
                if 'find_text' not in parameters:
                    raise ValidationError(
                        f"Action {index}: find_and_replace operation requires 'find_text' parameter"
                    )
                if 'replace_text' not in parameters:
                    raise ValidationError(
                        f"Action {index}: find_and_replace operation requires 'replace_text' parameter"
                    )
                # Validate find_text is not empty
                if not parameters['find_text']:
                    raise ValidationError(
                        f"Action {index}: 'find_text' cannot be empty"
                    )
        
        # Word-specific validations
        elif tool == 'word':
            if operation == 'create':
                if 'elements' in parameters:
                    if not isinstance(parameters['elements'], list):
                        raise ValidationError(
                            f"Action {index}: 'elements' must be a list"
                        )
            
            elif operation == 'format':
                if 'formatting' not in parameters:
                    raise ValidationError(
                        f"Action {index}: format operation requires 'formatting' parameter"
                    )
        
        # PDF-specific validations
        elif tool == 'pdf':
            if operation == 'merge':
                if 'file_paths' not in parameters:
                    raise ValidationError(
                        f"Action {index}: merge operation requires 'file_paths' parameter"
                    )
                if not isinstance(parameters['file_paths'], list):
                    raise ValidationError(
                        f"Action {index}: 'file_paths' must be a list"
                    )
                if len(parameters['file_paths']) < 2:
                    raise ValidationError(
                        f"Action {index}: merge requires at least 2 files"
                    )
    
    @staticmethod
    def extract_actions(parsed_response: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Extract the list of actions from a parsed response.
        
        Args:
            parsed_response: Validated parsed response
        
        Returns:
            List of action dictionaries
        """
        return parsed_response.get('actions', [])
    
    @staticmethod
    def extract_explanation(parsed_response: Dict[str, Any]) -> str:
        """Extract the explanation from a parsed response.
        
        Args:
            parsed_response: Validated parsed response
        
        Returns:
            Explanation string, or empty string if not present
        """
        return parsed_response.get('explanation', '')
