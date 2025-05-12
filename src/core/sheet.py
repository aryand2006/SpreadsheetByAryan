class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}

    def add_cell(self, cell_id, cell):
        self.cells[cell_id] = cell

    def remove_cell(self, cell_id):
        if cell_id in self.cells:
            del self.cells[cell_id]

    def get_cell(self, cell_id):
        return self.cells.get(cell_id, None)

    def get_all_cells(self):
        return self.cells.items()