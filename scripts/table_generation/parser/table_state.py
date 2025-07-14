class Span:
    def __init__(self, pos1, pos2, text):
        self.pos1 = pos1
        self.pos2 = pos2
        self.text = text

class TableState:
    def __init__(self):
        self._cur_i = 0
        self._cur_j = 0
        self.rows = 1
        self.cols = 1
        self.arr = [[None]]
        self.force_cutoffs = []
        self.spans = []

    def set_elm(self, elm):
        while self._cur_i >= self.rows:
            self._append_row()
        while self._cur_j >= self.cols:
            self._append_col()

        self.arr[self._cur_i][self._cur_j] = elm

    def reset_col(self):
        self._cur_j = 0

    def next_row(self):
        self._cur_i += 1

    def next_col(self):
        self._cur_j += 1

    def force_cutoff(self):
        self.force_cutoffs.append(self._cur_i)

    def add_span(self, text, length):
        extend_by = length - 1
        while self._cur_j + extend_by >= self.cols:
            self._append_col()

        span = Span((self._cur_i, self._cur_j), (self._cur_i, self._cur_j + extend_by), text)
        self.spans.append(span)

        self._cur_j += length

    def _append_row(self):
        # Add empty row to the end
        self.arr.append([None for _ in range(self.cols)])
        self.rows += 1

    def _append_col(self):
        # Add an empty element to each row
        for l in self.arr:
            l.append(None)
        self.cols += 1