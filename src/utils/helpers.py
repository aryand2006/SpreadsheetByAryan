import re

def calculate_percentage(part, whole):
    if whole == 0:
        return 0
    return (part / whole) * 100

def format_currency(value):
    return "${:,.2f}".format(value)

def parse_cell_reference(ref):
    """
    Parse a cell reference like 'A1' into column name and row number
    
    Args:
        ref (str): Cell reference (e.g. 'A1', 'BC123')
        
    Returns:
        tuple: (column_name, row_number)
    """
    # Regular expression to match a cell reference
    cell_ref_pattern = re.compile(r'([A-Z]+)(\d+)')
    match = cell_ref_pattern.match(ref)
    
    if not match:
        return None, None
    
    col_name, row = match.groups()
    try:
        row = int(row)
    except ValueError:
        return None, None
        
    return col_name, row

def column_name_to_index(col_name):
    """
    Convert column name (A, B, ..., Z, AA, AB, etc.) to 0-based index
    
    Args:
        col_name (str): Column name
        
    Returns:
        int: 0-based column index
    """
    index = 0
    for char in col_name:
        index = index * 26 + (ord(char.upper()) - ord('A') + 1)
    return index - 1

def index_to_column_name(index):
    """
    Convert 0-based column index to name (A, B, ..., Z, AA, AB, etc.)
    
    Args:
        index (int): 0-based column index
        
    Returns:
        str: Column name
    """
    if index < 0:
        return ""
        
    result = ""
    index += 1  # Convert to 1-based for calculation
    
    while index > 0:
        remainder = (index - 1) % 26
        result = chr(65 + remainder) + result
        index = (index - 1) // 26
        
    return result

def format_cell_address(row, col):
    """
    Format row and column indices as a cell address (e.g. A1)
    
    Args:
        row (int): 0-based row index
        col (int): 0-based column index
        
    Returns:
        str: Cell address (e.g. A1)
    """
    col_name = index_to_column_name(col)
    return f"{col_name}{row + 1}"  # Add 1 to convert to 1-based

def validate_formula(formula):
    # Basic validation for a formula string
    allowed_chars = set("0123456789+-*/()ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz ")
    return all(char in allowed_chars for char in formula)

def transpose_matrix(matrix):
    return [list(row) for row in zip(*matrix)]