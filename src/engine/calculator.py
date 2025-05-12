import re
import datetime
import statistics
import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta  
from math import ceil, floor, sqrt, sin, cos, tan, log, log10, exp, pi
from ..utils.helpers import parse_cell_reference

class Calculator:
    def __init__(self, sheet_view=None):
        self.sheet_view = sheet_view
        self.formula_pattern = re.compile(r'=([A-Z]+\([^)]*\)|[^=]*)')
        self.cell_ref_pattern = re.compile(r'([A-Z]+)(\d+)')
        self.operators = {
            '+': lambda x, y: x + y,
            '-': lambda x, y: x - y,
            '*': lambda x, y: x * y,
            '/': lambda x, y: x / y if y != 0 else float('inf'),
            '^': lambda x, y: x ** y,
            '%': lambda x, y: x % y,
        }
        self.formula_cache = {}  # Cache for formula results
        self.dependent_cells = {}  # Track cell dependencies
        # Expanded functions dictionary with Google Sheets-like functionality
        self.functions = self._initialize_functions()
        
    def _initialize_functions(self):
        """Initialize the full set of supported functions"""
        return {
            # Basic math functions
            'SUM': sum,
            'AVERAGE': lambda values: sum(values) / len(values) if values else 0,
            'COUNT': len,
            'MIN': lambda values: min(values) if values else 0,
            'MAX': lambda values: max(values) if values else 0,
            'PRODUCT': lambda values: np.prod(values) if values else 0,
            
            # Advanced math functions
            'ROUND': lambda values: round(values[0], int(values[1])) if len(values) >= 2 else round(values[0]),
            'ROUNDUP': lambda values: ceil(values[0] * 10**values[1]) / 10**values[1] if len(values) >= 2 else ceil(values[0]),
            'ROUNDDOWN': lambda values: floor(values[0] * 10**values[1]) / 10**values[1] if len(values) >= 2 else floor(values[0]),
            'ABS': lambda values: abs(values[0]) if values else 0,
            'SQRT': lambda values: sqrt(values[0]) if values and values[0] >= 0 else '#ERROR',
            'POWER': lambda values: values[0] ** values[1] if len(values) >= 2 else '#ERROR',
            'MOD': lambda values: values[0] % values[1] if len(values) >= 2 else '#ERROR',
            'GCD': lambda values: np.gcd.reduce(np.array(values, dtype=int)) if values else '#ERROR',
            'LCM': lambda values: np.lcm.reduce(np.array(values, dtype=int)) if values else '#ERROR',
            'FACT': lambda values: np.math.factorial(int(values[0])) if values else '#ERROR',
            'RAND': lambda values: np.random.random(),
            'RANDBETWEEN': lambda values: np.random.randint(values[0], values[1]+1) if len(values) >= 2 else '#ERROR',
            'PI': lambda values: pi,
            'SIN': lambda values: sin(values[0]) if values else '#ERROR',
            'COS': lambda values: cos(values[0]) if values else '#ERROR',
            'TAN': lambda values: tan(values[0]) if values else '#ERROR',
            'LN': lambda values: log(values[0]) if values and values[0] > 0 else '#ERROR',
            'LOG10': lambda values: log10(values[0]) if values and values[0] > 0 else '#ERROR',
            'LOG': lambda values: log(values[0], values[1]) if len(values) >= 2 and values[0] > 0 and values[1] > 0 else '#ERROR',
            'EXP': lambda values: exp(values[0]) if values else '#ERROR',
            
            # Statistical functions
            'STDEV': lambda values: statistics.stdev(values) if len(values) > 1 else '#ERROR',
            'STDEVP': lambda values: statistics.pstdev(values) if values else '#ERROR',
            'VAR': lambda values: statistics.variance(values) if len(values) > 1 else '#ERROR',
            'VARP': lambda values: statistics.pvariance(values) if values else '#ERROR',
            'MEDIAN': lambda values: statistics.median(values) if values else '#ERROR',
            'MODE': lambda values: statistics.mode(values) if values else '#ERROR',
            'PERCENTILE': lambda values: np.percentile(values[:-1], values[-1]*100) if len(values) >= 2 else '#ERROR',
            'QUARTILE': lambda values: np.percentile(values[:-1], values[-1]*25) if len(values) >= 2 else '#ERROR',
            'CORREL': lambda values: np.corrcoef(values[:len(values)//2], values[len(values)//2:])[0,1] if len(values) >= 2 and len(values) % 2 == 0 else '#ERROR',
            
            # Logical functions
            'IF': self.if_function,
            'AND': lambda values: all(bool(v) for v in values),
            'OR': lambda values: any(bool(v) for v in values),
            'NOT': lambda values: not bool(values[0]) if values else '#ERROR',
            'XOR': lambda values: bool(sum(bool(v) for v in values) % 2) if values else '#ERROR',
            'TRUE': lambda values: True,
            'FALSE': lambda values: False,
            'ISBLANK': lambda values: not values or not values[0],
            'ISERROR': lambda values: isinstance(values[0], str) and values[0].startswith('#') if values else '#ERROR',
            'ISLOGICAL': lambda values: isinstance(values[0], bool) if values else '#ERROR',
            'ISNA': lambda values: values[0] == '#N/A' if values else '#ERROR',
            'ISTEXT': lambda values: isinstance(values[0], str) and not values[0].startswith('#') if values else '#ERROR',
            'ISNUMBER': lambda values: isinstance(values[0], (int, float)) if values else '#ERROR',
            
            # Text functions
            'CONCATENATE': lambda values: ''.join(str(v) for v in values),
            'LEFT': lambda values: str(values[0])[:int(values[1])] if len(values) >= 2 else '#ERROR',
            'RIGHT': lambda values: str(values[0])[-int(values[1]):] if len(values) >= 2 else '#ERROR',
            'MID': lambda values: str(values[0])[int(values[1])-1:int(values[1])+int(values[2])-1] if len(values) >= 3 else '#ERROR',
            'LEN': lambda values: len(str(values[0])) if values else 0,
            'LOWER': lambda values: str(values[0]).lower() if values else '',
            'UPPER': lambda values: str(values[0]).upper() if values else '',
            'PROPER': lambda values: str(values[0]).title() if values else '',
            'TRIM': lambda values: str(values[0]).strip() if values else '',
            'REPLACE': lambda values: str(values[0])[:int(values[1])-1] + str(values[3]) + str(values[0])[int(values[1])+int(values[2])-1:] if len(values) >= 4 else '#ERROR',
            'SUBSTITUTE': lambda values: str(values[0]).replace(str(values[1]), str(values[2]), int(values[3]) if len(values) >= 4 else -1) if len(values) >= 3 else '#ERROR',
            'REPT': lambda values: str(values[0]) * int(values[1]) if len(values) >= 2 else '#ERROR',
            'FIND': lambda values: str(values[0]).find(str(values[1]), int(values[2])-1 if len(values) >= 3 else 0)+1 if len(values) >= 2 else '#ERROR',
            'SEARCH': lambda values: str(values[0]).lower().find(str(values[1]).lower(), int(values[2])-1 if len(values) >= 3 else 0)+1 if len(values) >= 2 else '#ERROR',
            'CHAR': lambda values: chr(int(values[0])) if values else '#ERROR',
            'CODE': lambda values: ord(str(values[0])[0]) if values and str(values[0]) else '#ERROR',
            'EXACT': lambda values: str(values[0]) == str(values[1]) if len(values) >= 2 else '#ERROR',
            'TEXT': lambda values: self.format_text(values[0], values[1]) if len(values) >= 2 else '#ERROR',
            'DOLLAR': lambda values: self.format_currency(values[0], int(values[1]) if len(values) >= 2 else 2) if values else '#ERROR',
            
            # Date and time functions
            'NOW': lambda values: datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'TODAY': lambda values: datetime.date.today().strftime('%Y-%m-%d'),
            'DATE': lambda values: self.make_date(values[0], values[1], values[2]) if len(values) >= 3 else '#ERROR',
            'TIME': lambda values: self.make_time(values[0], values[1], values[2]) if len(values) >= 3 else '#ERROR',
            'DAY': lambda values: self.parse_date(values[0]).day if values else '#ERROR',
            'MONTH': lambda values: self.parse_date(values[0]).month if values else '#ERROR',
            'YEAR': lambda values: self.parse_date(values[0]).year if values else '#ERROR',
            'WEEKDAY': lambda values: self.parse_date(values[0]).weekday() + (1 if len(values) < 2 or values[1] == 1 else 0) if values else '#ERROR',
            'HOUR': lambda values: self.parse_time(values[0]).hour if values else '#ERROR',
            'MINUTE': lambda values: self.parse_time(values[0]).minute if values else '#ERROR',
            'SECOND': lambda values: self.parse_time(values[0]).second if values else '#ERROR',
            'EDATE': lambda values: (self.parse_date(values[0]) + relativedelta(months=int(values[1]))).strftime('%Y-%m-%d') if len(values) >= 2 else '#ERROR',
            'EOMONTH': lambda values: (self.parse_date(values[0]) + relativedelta(months=int(values[1]), day=31)).strftime('%Y-%m-%d') if len(values) >= 2 else '#ERROR',
            'NETWORKDAYS': lambda values: self.calculate_workdays(values[0], values[1]) if len(values) >= 2 else '#ERROR',
            'WORKDAY': lambda values: self.add_workdays(values[0], int(values[1])).strftime('%Y-%m-%d') if len(values) >= 2 else '#ERROR',
            'DATEDIF': lambda values: self.date_diff(values[0], values[1], values[2]) if len(values) >= 3 else '#ERROR',
            
            # Lookup and reference functions
            'VLOOKUP': self.vlookup_function,
            'HLOOKUP': self.hlookup_function,
            'INDEX': self.index_function,
            'MATCH': self.match_function,
            'ADDRESS': lambda values: self.get_address(values[0], values[1], values[2] if len(values) >= 3 else 1) if len(values) >= 2 else '#ERROR',
            'INDIRECT': self.indirect_function,
            'ROW': lambda values: int(values[0].split(':')[0][1:]) if values and ':' in values[0] else self.current_row + 1 if not values else '#ERROR',
            'COLUMN': lambda values: self.column_name_to_index(values[0].split(':')[0][0]) + 1 if values and ':' in values[0] else self.current_col + 1 if not values else '#ERROR',
            'ROWS': lambda values: len(self.get_range_values(values[0], self.current_row, self.current_col)) if values else '#ERROR',
            'COLUMNS': lambda values: len(self.get_range_values(values[0], self.current_row, self.current_col)[0]) if values and self.get_range_values(values[0], self.current_row, self.current_col) else '#ERROR',
            'OFFSET': self.offset_function,
            
            # Financial functions
            'PMT': lambda values: self.pmt_function(values[0], values[1], values[2], values[3] if len(values) >= 4 else 0, values[4] if len(values) >= 5 else 0) if len(values) >= 3 else '#ERROR',
            'FV': lambda values: self.fv_function(values[0], values[1], values[2], values[3] if len(values) >= 4 else 0, values[4] if len(values) >= 5 else 0) if len(values) >= 3 else '#ERROR',
            'PV': lambda values: self.pv_function(values[0], values[1], values[2], values[3] if len(values) >= 4 else 0, values[4] if len(values) >= 5 else 0) if len(values) >= 3 else '#ERROR',
            'NPV': lambda values: self.npv_function(values[0], values[1:]) if len(values) >= 2 else '#ERROR',
            'IRR': lambda values: self.irr_function(values) if values else '#ERROR',
            'RATE': lambda values: self.rate_function(values[0], values[1], values[2], values[3] if len(values) >= 4 else 0, values[4] if len(values) >= 5 else 0, values[5] if len(values) >= 6 else 0.1) if len(values) >= 3 else '#ERROR',
            'NPER': lambda values: self.nper_function(values[0], values[1], values[2], values[3] if len(values) >= 4 else 0, values[4] if len(values) >= 5 else 0) if len(values) >= 3 else '#ERROR',
            'SLN': lambda values: (values[0] - values[1]) / values[2] if len(values) >= 3 and values[2] != 0 else '#ERROR',
            'SYD': lambda values: self.syd_function(values[0], values[1], values[2], values[3]) if len(values) >= 4 else '#ERROR',
            
            # Array functions
            'TRANSPOSE': lambda values: np.transpose(values).flatten().tolist() if values else '#ERROR',
            'MMULT': lambda values: np.matmul(np.array(values[:len(values)//2]).reshape(-1, int(sqrt(len(values)//2))), 
                                           np.array(values[len(values)//2:]).reshape(-1, int(sqrt(len(values)//2)))).flatten().tolist() if len(values) >= 4 else '#ERROR',
            'SUMPRODUCT': lambda values: sum(values[i] * values[i+len(values)//2] for i in range(len(values)//2)) if len(values) >= 2 and len(values) % 2 == 0 else '#ERROR',
            
            # Information functions
            'ISEVEN': lambda values: int(values[0]) % 2 == 0 if values else '#ERROR',
            'ISODD': lambda values: int(values[0]) % 2 == 1 if values else '#ERROR',
            'NA': lambda values: '#N/A',
            'ERROR.TYPE': lambda values: {'#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7}.get(values[0], '#N/A') if values else '#ERROR',
            'INFO': self.info_function,
            
            # Web functions
            'ENCODEURL': lambda values: self.url_encode(values[0]) if values else '#ERROR',
            'HYPERLINK': self.hyperlink_function,
            
            # Advanced Google Sheets specific functions
            'GOOGLEFINANCE': lambda values: '#FEATURE_NOT_IMPLEMENTED',
            'IMPORTDATA': lambda values: '#FEATURE_NOT_IMPLEMENTED',
            'IMPORTRANGE': lambda values: '#FEATURE_NOT_IMPLEMENTED',
            'IMAGE': lambda values: '#FEATURE_NOT_IMPLEMENTED',
        }
    
    def set_sheet_view(self, sheet_view):
        """Set the sheet view to work with"""
        self.sheet_view = sheet_view

    def evaluate(self, formula, row, col):
        """Evaluate a formula in the context of the sheet"""
        if not formula.startswith('='):
            return formula  # Not a formula
            
        # Extract the expression part (remove the leading '=')
        expression = formula[1:].strip()
        
        try:
            # Check if it's a function call
            function_match = re.match(r'([A-Z]+)\((.*)\)', expression)
            if function_match:
                func_name = function_match.group(1)
                args_str = function_match.group(2)
                
                if func_name in self.functions:
                    # Parse arguments
                    values = self.parse_function_arguments(args_str, row, col)
                    # Apply the function
                    result = self.functions[func_name](values)
                    return str(result)
                else:
                    return f"#ERROR: Unknown function {func_name}"
            
            # Otherwise evaluate as an expression with cell references
            result = self.evaluate_expression(expression, row, col)
            return str(result)
        except Exception as e:
            return f"#ERROR: {str(e)}"

    def parse_function_arguments(self, args_str, current_row, current_col):
        """Parse function arguments into values"""
        values = []
        
        # Split by commas, but respect ranges like A1:B5
        args = []
        current_arg = ""
        in_range = False
        
        for char in args_str:
            if char == ',' and not in_range:
                args.append(current_arg.strip())
                current_arg = ""
            elif char == ':':
                in_range = True
                current_arg += char
            else:
                current_arg += char
        
        if current_arg:
            args.append(current_arg.strip())
        
        # Process each argument
        for arg in args:
            if ':' in arg:  # Range like A1:B5
                range_values = self.get_range_values(arg, current_row, current_col)
                values.extend(range_values)
            else:  # Single cell or value
                try:
                    # Try as a number
                    values.append(float(arg))
                except ValueError:
                    # Try as a cell reference
                    cell_value = self.get_cell_value(arg, current_row, current_col)
                    if cell_value is not None:
                        try:
                            values.append(float(cell_value))
                        except ValueError:
                            # Non-numeric cell value, skip it
                            pass
        
        return values

    def get_range_values(self, range_str, current_row, current_col):
        """Get all values in a range like A1:B5"""
        if not self.sheet_view:
            return []
            
        range_parts = range_str.split(':')
        if len(range_parts) != 2:
            raise ValueError(f"Invalid range format: {range_str}")
            
        start_ref = range_parts[0].strip()
        end_ref = range_parts[1].strip()
        
        # Parse cell references
        try:
            match_start = self.cell_ref_pattern.match(start_ref)
            match_end = self.cell_ref_pattern.match(end_ref)
            
            if not match_start or not match_end:
                return []
                
            start_col_name, start_row_str = match_start.groups()
            end_col_name, end_row_str = match_end.groups()
            
            # Convert to indices
            start_row = int(start_row_str) - 1  # Convert to 0-based
            end_row = int(end_row_str) - 1
            
            # Convert column letters to indices
            start_col = self.column_name_to_index(start_col_name)
            end_col = self.column_name_to_index(end_col_name)
        except Exception as e:
            print(f"Error parsing cell references: {e}")
            return []
        
        # Make sure start is less than end for proper range
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col
        
        values = []
        try:
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    if (row < 0 or col < 0 or 
                        row >= self.sheet_view.rowCount() or 
                        col >= self.sheet_view.columnCount()):
                        continue
                    
                    item = self.sheet_view.item(row, col)
                    if not item:
                        # Empty cell counts as 0 for math functions
                        values.append(0)
                        continue
                        
                    cell_value = item.text()
                    
                    if not cell_value:  # Empty string
                        values.append(0)
                        continue
                        
                    # Try to convert to number for calculation
                    try:
                        values.append(float(cell_value))
                    except ValueError:
                        # For non-numeric values in numeric functions, use 0
                        if not cell_value.startswith('='):
                            values.append(0) 
        except Exception as e:
            print(f"Error processing range {range_str}: {e}")
        
        return values

    def get_cell_value(self, cell_ref, current_row, current_col):
        """Get a cell's value from its reference"""
        # Check if sheet_view is set
        if not self.sheet_view:
            return "#ERROR: No sheet available"
            
        match = self.cell_ref_pattern.match(cell_ref)
        if not match:
            return None
            
        col_name, row = match.groups()
        try:
            row = int(row) - 1  # Convert to 0-based
        except ValueError:
            return "#ERROR: Invalid cell reference"
        
        # Convert column name to index
        col = 0
        for char in col_name:
            col = col * 26 + (ord(char) - ord('A') + 1)
        col -= 1  # Convert to 0-based
        
        try:
            # Check if the row and column are within sheet bounds
            if (row < 0 or col < 0 or 
                row >= self.sheet_view.table.rowCount() or 
                col >= self.sheet_view.table.columnCount()):
                return "#ERROR: Cell reference out of bounds"
                
            return self.sheet_view.get_cell_value(row, col)
        except AttributeError:
            # This will catch if sheet_view.table or sheet_view.get_cell_value doesn't exist
            return "#ERROR: Unable to access sheet data"
        except Exception as e:
            return f"#ERROR: {str(e)}"

    def evaluate_expression(self, expression, current_row, current_col):
        """Evaluate an expression with cell references and named ranges"""
        # First, check for named ranges and replace them
        if hasattr(self, 'main_window') and hasattr(self.main_window, 'named_ranges'):
            for name, range_data in self.main_window.named_ranges.items():
                # Replace named ranges with their actual cell references
                expression = expression.replace(name, range_data['range'])
        
        # Then replace cell references with their values
        def replace_cell_ref(match):
            cell_ref = match.group(0)
            cell_value = self.get_cell_value(cell_ref, current_row, current_col)
            if cell_value is None or cell_value == "":
                return '0'
            try:
                return str(float(cell_value))
            except ValueError:
                return '0'  # Non-numeric cell values treated as 0
        
        # Replace all cell references with their values
        numeric_expr = self.cell_ref_pattern.sub(replace_cell_ref, expression)
        
        # Safely evaluate the expression
        try:
            return eval(numeric_expr)
        except Exception as e:
            raise ValueError(f"Invalid expression: {expression}")

    def validate_formula(self, formula):
        """Validate a formula's syntax"""
        if not formula.startswith('='):
            return True  # Not a formula, so no validation needed
            
        expression = formula[1:].strip()
        
        # Check for function call
        function_match = re.match(r'([A-Z]+)\((.*)\)', expression)
        if function_match:
            func_name = function_match.group(1)
            if func_name not in self.functions:
                return False  # Unknown function
            
        # Check for balanced parentheses
        if expression.count('(') != expression.count(')'):
            return False
            
        # More validation can be added as needed
        
        return True

    def if_function(self, values):
        """Implements Excel's IF function: IF(condition, value_if_true, value_if_false)"""
        if len(values) < 3:
            return '#ERROR: IF requires 3 arguments'
        return values[1] if bool(values[0]) else values[2]
    
    def vlookup_function(self, values):
        """Implements Excel's VLOOKUP function"""
        if len(values) < 3:
            return '#ERROR: VLOOKUP requires at least 3 arguments'
        
        lookup_value = values[0]
        lookup_range = values[1]  # This should be a range reference
        col_index = int(values[2])
        exact_match = len(values) < 4 or bool(values[3])
        
        # Implementation would need access to the referenced data
        # This is a simplified placeholder
        return '#FEATURE_NOT_IMPLEMENTED'
    
    def hlookup_function(self, values):
        """Implements Excel's HLOOKUP function"""
        if len(values) < 3:
            return '#ERROR: HLOOKUP requires at least 3 arguments'
        
        # Similar to VLOOKUP but for horizontal lookups
        return '#FEATURE_NOT_IMPLEMENTED'
    
    def index_function(self, values):
        """Implements Excel's INDEX function"""
        if len(values) < 3:
            return '#ERROR: INDEX requires at least 3 arguments'
            
        # Implementation for array reference, row number, column number
        return '#FEATURE_NOT_IMPLEMENTED'
    
    def match_function(self, values):
        """Implements Excel's MATCH function"""
        if len(values) < 2:
            return '#ERROR: MATCH requires at least 2 arguments'
            
        # Implementation for lookup value, lookup array, match type
        return '#FEATURE_NOT_IMPLEMENTED'
    
    def parse_date(self, date_str):
        """Parse a date string into a datetime object"""
        try:
            # Try various date formats
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y', '%d/%m/%Y']:
                try:
                    return datetime.datetime.strptime(str(date_str), fmt)
                except ValueError:
                    continue
            # If none of the formats work
            return datetime.datetime.now()  # Default to current date
        except:
            return datetime.datetime.now()

    def parse_cell_reference(self, ref, current_row, current_col):
        """Parse cell reference, handling absolute references like $A$1"""
        if not ref:
            return None, None
            
        match = re.match(r'(\$?)([A-Z]+)(\$?)(\d+)', ref)
        if not match:
            return None, None
        
        col_abs, col_name, row_abs, row_num = match.groups()
        
        col_index = 0
        for char in col_name:
            col_index = col_index * 26 + (ord(char.upper()) - ord('A') + 1)
        col_index -= 1
        
        row_index = int(row_num) - 1
        
        if not col_abs:
            col_index = col_index + current_col
        
        if not row_abs:
            row_index = row_index + current_row
        
        return row_index, col_index

    def format_currency(self, value, decimal_places=2):
        """Format a value as currency with the specified number of decimal places"""
        try:
            return f"${float(value):,.{decimal_places}f}"
        except (ValueError, TypeError):
            return '#ERROR: Invalid value for currency formatting'

    def make_date(self, year, month, day):
        """Create a date from year, month, day values"""
        try:
            return datetime.date(int(year), int(month), int(day)).strftime('%Y-%m-%d')
        except ValueError:
            return '#ERROR: Invalid date'

    def make_time(self, hour, minute, second):
        """Create a time from hour, minute, second values"""
        try:
            return datetime.time(int(hour), int(minute), int(second)).strftime('%H:%M:%S')
        except ValueError:
            return '#ERROR: Invalid time'
    
    def format_text(self, value, format_string):
        """Format a value according to a format string (simplified)"""
        try:
            if format_string == '0.00%':
                return f"{float(value)*100:.2f}%"
            elif format_string == '0.00':
                return f"{float(value):.2f}"
            elif format_string == '#,##0':
                return f"{int(float(value)):,}"
            elif format_string == '#,##0.00':
                return f"{float(value):,.2f}"
            elif format_string.startswith('$'):
                decimals = format_string.count('0')
                return f"${float(value):,.{decimals}f}"
            else:
                return str(value)
        except:
            return '#ERROR: Invalid formatting'
    
    def calculate_workdays(self, start_date, end_date):
        """Calculate number of workdays between two dates"""
        try:
            start = self.parse_date(start_date)
            end = self.parse_date(end_date)
            days = (end - start).days + 1
            weeks = days // 7
            extra_days = days % 7
            weekend_days = 0
            
            # Count weekend days in the whole weeks
            weekend_days = weeks * 2
            
            # Count weekend days in the extra days
            start_weekday = start.weekday()
            for i in range(extra_days):
                day = (start_weekday + i) % 7
                if day >= 5:  # 5=Saturday, 6=Sunday
                    weekend_days += 1
            
            return days - weekend_days
        except:
            return '#ERROR: Invalid date range'

    def add_workdays(self, start_date, days):
        """Add a specified number of workdays to a date"""
        try:
            start = self.parse_date(start_date)
            current_date = start
            remaining_days = days
            
            while remaining_days > 0:
                current_date += datetime.timedelta(days=1)
                if current_date.weekday() < 5:  # 0-4 are weekdays
                    remaining_days -= 1
            
            return current_date
        except:
            return '#ERROR: Invalid date or days'
    
    def date_diff(self, start_date, end_date, unit):
        """Calculate difference between dates based on unit"""
        try:
            start = self.parse_date(start_date)
            end = self.parse_date(end_date)
            
            if unit == 'Y':  # Years
                return end.year - start.year
            elif unit == 'M':  # Months
                return (end.year - start.year) * 12 + (end.month - start.month)
            elif unit == 'D':  # Days
                return (end - start).days
            else:
                return '#ERROR: Invalid unit'
        except:
            return '#ERROR: Invalid date range'

    def hyperlink_function(self, values):
        """Create hyperlink with URL and optional display text"""
        if len(values) < 1:
            return '#ERROR: HYPERLINK requires at least 1 argument'
        
        url = str(values[0])
        display_text = str(values[1]) if len(values) >= 2 else url
        
        # In Google Sheets, this would create a clickable link
        # Here we'll return the display text with the URL in parentheses
        return f"{display_text} ({url})"
    
    def url_encode(self, text):
        """URL encode a string"""
        from urllib.parse import quote
        try:
            return quote(str(text))
        except:
            return '#ERROR: URL encoding failed'
    
    def info_function(self, values):
        """Return system information based on type"""
        if not values:
            return '#ERROR: INFO requires 1 argument'
            
        info_type = values[0].upper() if isinstance(values[0], str) else str(values[0])
        
        if info_type == "DIRECTORY":
            import os
            return os.getcwd()
        elif info_type == "NUMFILE":
            return "1"  # Only current file is open in this app
        elif info_type == "OSVERSION":
            import platform
            return platform.platform()
        elif info_type == "RECALC":
            return "Automatic"
        elif info_type == "RELEASE":
            return "PySpreadsheet 1.0"
        else:
            return '#ERROR: Unknown INFO type'

    def pmt_function(self, rate, nper, pv, fv=0, type=0):
        """Calculate the payment for a loan"""
        try:
            rate = float(rate)
            nper = float(nper)
            pv = float(pv)
            fv = float(fv)
            type = float(type)
            
            if rate == 0:
                return -(pv + fv)/nper
                
            pvif = (1 + rate)**nper
            pmt = rate / (pvif - 1) * -(pv * pvif + fv)
            
            if type == 1:
                pmt /= (1 + rate)
                
            return pmt
        except:
            return '#ERROR: Invalid parameters for PMT'
    
    def fv_function(self, rate, nper, pmt, pv=0, type=0):
        """Calculate future value of an investment"""
        try:
            rate = float(rate)
            nper = float(nper)
            pmt = float(pmt)
            pv = float(pv)
            type = float(type)
            
            if rate == 0:
                return -(pv + pmt * nper)
                
            pvif = (1 + rate)**nper
            fv = -(pv * pvif + pmt * (1 + rate * type) * (pvif - 1) / rate)
                
            return fv
        except:
            return '#ERROR: Invalid parameters for FV'
    
    def pv_function(self, rate, nper, pmt, fv=0, type=0):
        """Calculate present value of an investment"""
        try:
            rate = float(rate)
            nper = float(nper)
            pmt = float(pmt)
            fv = float(fv)
            type = float(type)
            
            if rate == 0:
                return -(fv + pmt * nper)
            
            pvif = (1 + rate)**nper
            pv = -( fv + pmt * (1 + rate * type) * (pvif - 1) / rate ) / pvif
                
            return pv
        except:
            return '#ERROR: Invalid parameters for PV'

    def recalculate_all(self):
        """Recalculate all formulas in the sheet"""
        if not self.sheet_view:
            return
            
        # Clear cache
        self.formula_cache = {}
        
        # Get all cells with formulas
        formula_cells = []
        for row in range(self.sheet_view.table.rowCount()):
            for col in range(self.sheet_view.table.columnCount()):
                value = self.sheet_view.get_cell_value(row, col)
                if value and value.startswith('='):
                    formula_cells.append((row, col, value))
        
        # Build dependency graph
        self.build_dependency_graph(formula_cells)
        
        # Calculate in dependency order
        evaluated_cells = set()
        
        # Helper function for topological sort
        def evaluate_with_dependencies(row, col, formula):
            cell_key = f"{row},{col}"
            
            # Skip if already evaluated
            if cell_key in evaluated_cells:
                return
                
            # Evaluate dependencies first
            if cell_key in self.dependent_cells:
                for dep_row, dep_col in self.dependent_cells[cell_key]:
                    dep_formula = self.sheet_view.get_cell_value(dep_row, dep_col)
                    if dep_formula and dep_formula.startswith('='):
                        evaluate_with_dependencies(dep_row, dep_col, dep_formula)
            
            # Now evaluate this cell
            result = self.evaluate(formula, row, col)
            self.sheet_view.set_cell_display_value(row, col, result)
            evaluated_cells.add(cell_key)
        
        # Evaluate all formulas in dependency order
        for row, col, formula in formula_cells:
            evaluate_with_dependencies(row, col, formula)
    
    def build_dependency_graph(self, formula_cells):
        """Build a graph of cell dependencies"""
        self.dependent_cells = {}
        
        for row, col, formula in formula_cells:
            cell_key = f"{row},{col}"
            
            # Extract cell references from the formula
            refs = self.extract_cell_references(formula)
            
            # Add dependencies
            for ref_row, ref_col in refs:
                dep_key = f"{ref_row},{ref_col}"
                if dep_key not in self.dependent_cells:
                    self.dependent_cells[dep_key] = []
                self.dependent_cells[dep_key].append((row, col))
    
    def extract_cell_references(self, formula):
        """Extract cell references from a formula"""
        refs = []
        
        # This is a simplified extractor - a real one would parse the formula properly
        for match in re.finditer(r'[A-Z]+\d+', formula):
            cell_ref = match.group(0)
            col_name = ''.join(filter(str.isalpha, cell_ref))
            row_num = int(''.join(filter(str.isdigit, cell_ref)))
            
            # Convert to 0-based indices
            col_idx = self.column_name_to_index(col_name)
            row_idx = row_num - 1
            
            refs.append((row_idx, col_idx))
            
        return refs
    
    def column_name_to_index(self, name):
        """Convert column name (A, B, AA, etc.) to index"""
        index = 0
        for char in name:
            index = index * 26 + (ord(char.upper()) - ord('A') + 1)
        return index - 1
    
    # Add these missing methods referenced in _initialize_functions
    def indirect_function(self, values):
        """Implements Excel's INDIRECT function - returns the reference specified by a text string"""
        if not values:
            return '#ERROR: INDIRECT requires a reference as text'
            
        ref_text = str(values[0])
        try:
            # Attempt to get the cell value from the reference
            if ':' in ref_text:
                # Range reference
                return self.get_range_values(ref_text, self.current_row, self.current_col)
            else:
                # Single cell reference
                return self.get_cell_value(ref_text, self.current_row, self.current_col)
        except Exception as e:
            return f'#ERROR: {str(e)}'
    
    def offset_function(self, values):
        """Implements Excel's OFFSET function - returns a reference offset from a given reference"""
        if len(values) < 3:
            return '#ERROR: OFFSET requires at least 3 arguments'
            
        # Implementation would need reference, rows, cols, [height], [width]
        return '#FEATURE_NOT_IMPLEMENTED'
            
    def npv_function(self, rate, values):
        """Calculate Net Present Value for a series of cash flows"""
        try:
            rate = float(rate)
            npv = 0
            for i, value in enumerate(values):
                npv += float(value) / ((1 + rate) ** (i + 1))
            return npv
        except:
            return '#ERROR: Invalid parameters for NPV'
    
    def irr_function(self, values):
        """Calculate Internal Rate of Return for a series of cash flows"""
        try:
            import numpy as np
            # Use numpy's IRR function
            return np.irr(values)
        except:
            return '#ERROR: Invalid cash flow sequence'
    
    def rate_function(self, nper, pmt, pv, fv=0, type=0, guess=0.1):
        """Calculate the interest rate per period"""
        try:
            # Simplified implementation - would need Newton's method for accurate results
            return '#FEATURE_NOT_IMPLEMENTED'
        except:
            return '#ERROR: Invalid parameters for RATE'
    
    def nper_function(self, rate, pmt, pv, fv=0, type=0):
        """Calculate the number of periods for an investment"""
        try:
            rate = float(rate)
            pmt = float(pmt)
            pv = float(pv)
            fv = float(fv)
            type = float(type)
            
            if rate == 0:
                return -(fv + pv) / pmt
                
            if pmt == 0:
                return float('inf')  # Cannot calculate if no payments
                
            # Calculate using the formula
            return np.log((pmt * (1 + rate * type) - fv * rate) / 
                         (pmt * (1 + rate * type) + pv * rate)) / np.log(1 + rate)
        except:
            return '#ERROR: Invalid parameters for NPER'
    
    def syd_function(self, cost, salvage, life, period):
        """Calculate Sum of Years' Digits Depreciation"""
        try:
            cost = float(cost)
            salvage = float(salvage)
            life = int(life)
            period = int(period)
            
            if life <= 0 or period <= 0 or period > life:
                return '#ERROR: Invalid life or period'
                
            # Calculate sum of years' digits
            sum_of_years = life * (life + 1) / 2
            
            # Calculate depreciation
            return ((life - period + 1) / sum_of_years) * (cost - salvage)
        except:
            return '#ERROR: Invalid parameters for SYD'