class TableState:
    def __init__(self):
        self._cur_i = 0
        self._cur_j = 0
        self.rows = 1
        self.cols = 1
        self._arr = [[None]]

    def set_elm(self, elm):
        while self._cur_i >= self.rows:
            self._append_row()
        while self._cur_j >= self.cols:
            self._append_col()

        self._arr[self._cur_i][self._cur_j] = elm

    def next_row(self):
        self._cur_i += 1

    def next_col(self):
        self._cur_j += 1

    def _append_row(self):
        # Add empty row to the end
        self._arr.append([None for _ in range(self.cols)])
        self.rows += 1

    def _append_col(self):
        # Add an empty element to each row
        for l in self._arr:
            l.append(None)
        self.cols += 1