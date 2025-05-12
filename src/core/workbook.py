class Workbook:
    def __init__(self):
        self.sheets = {}

    def add_sheet(self, sheet_name):
        if sheet_name not in self.sheets:
            self.sheets[sheet_name] = []
        else:
            raise ValueError(f"Sheet '{sheet_name}' already exists.")

    def remove_sheet(self, sheet_name):
        if sheet_name in self.sheets:
            del self.sheets[sheet_name]
        else:
            raise ValueError(f"Sheet '{sheet_name}' does not exist.")

    def get_sheet(self, sheet_name):
        return self.sheets.get(sheet_name, None)

    def save(self, file_path):
        # Logic to save the workbook to a file
        pass

    def load(self, file_path):
        # Logic to load the workbook from a file
        pass