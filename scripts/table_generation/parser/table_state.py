from dataclasses import dataclass, field
from typing import Tuple, List, Dict, Any

@dataclass
class Span:
    pos1 : Tuple[int, int]
    pos2 : Tuple[int, int]
    text : str

@dataclass
class Text:
    text : str = ""
    style : str = ""

class TableState:
    def __init__(self):
        self._cur_i = 0
        self._cur_j = 0
        self.rows = 1
        self.cols = 1
        self.arr : List[List[Text]] = [[Text()]]
        self.force_cutoffs = []
        self.spans = []

    def _expand(self):
        while self._cur_i >= self.rows:
            self._append_row()
        while self._cur_j >= self.cols:
            self._append_col()

    def set_text(self, text):
        self._expand()
        self.arr[self._cur_i][self._cur_j].text = text

    def set_style(self, style={}):
        self._expand()
        self.arr[self._cur_i][self._cur_j].style = style

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
        self.arr.append([Text() for _ in range(self.cols)])
        self.rows += 1

    def _append_col(self):
        # Add an empty element to each row
        for l in self.arr:
            l.append(Text())
        self.cols += 1