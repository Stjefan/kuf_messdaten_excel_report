import logging

class ExcelWorksheet:
    def __init__(self, no_rows, no_cols, wb, ws):
        self.ws_cells = {}
        for i in range(0, no_rows):
            for j in range(0, no_cols):
                self.ws_cells[(i, j)] = ExcelCell(i, j, wb, ws)

    def merge_cells(self, from_r, from_c, to_r, to_c):
        pop_all_but_one = True
        for i in [j for j in self.ws_cells if
                from_r <= j[0] <= to_r
                and to_c >= j[1] >= from_c
                ]:
            if pop_all_but_one:
                first = i
                pop_all_but_one = False
                self.ws_cells[i].format.set_text_wrap(True)
                self.ws_cells[i].format.set_align("center")
                # self.ws_cells[i].format.set_top(2)
                # self.ws_cells[i].format.set_left(2)
                # self.ws_cells[i].format.set_right(2)
                # self.ws_cells[i].format.set_bottom(2)
                self.ws_cells[i].is_merged = True
                self.ws_cells[i].merging_information = {
                    "from_r": from_r,
                    "from_c": from_c,
                    "to_r": to_r,
                    "to_c": to_c
                }
            else:
                if self.ws_cells[i].format.top != 0:
                    self.ws_cells[first].format.set_top(self.ws_cells[i].format.top)
                    logging.debug(f"found top at {i}")
                if self.ws_cells[i].format.left != 0:
                    logging.debug(f"found left at {i}")
                    self.ws_cells[first].format.set_left(self.ws_cells[i].format.left)
                if self.ws_cells[i].format.bottom != 0:
                    self.ws_cells[first].format.set_bottom(self.ws_cells[i].format.bottom)
                    logging.debug(f"found bottom at {i}")
                if self.ws_cells[i].format.right != 0:
                    logging.debug(f"found right at {i}")
                    self.ws_cells[first].format.set_right(self.ws_cells[i].format.right)
                self.ws_cells.pop(i)

    def make_box_around_cells(self, from_r, from_c, to_r, to_c, boxtype = 2, bg_color = None):
        for i in [j for j in self.ws_cells if
                from_r <= j[0] <= to_r
                and to_c >= j[1] >= from_c and j[1]
                ]:
            if bg_color is not None:
                self.ws_cells[i].format.set_bg_color(bg_color)
            pass
            # self.ws_cells[i].format.set_bg_color('red')
        for i in [j for j in self.ws_cells if
                from_r <= j[0] <= to_r
                and from_c == j[1]
                ]:
            self.ws_cells[i].format.set_left(boxtype)

        for i in [j for j in self.ws_cells if
                from_r <= j[0] <= to_r
                and to_c == j[1]
                ]:
            self.ws_cells[i].format.set_right(boxtype)

        for i in [j for j in self.ws_cells if
                from_r == j[0]
                and to_c >= j[1] >= from_c and j[1]
                ]:
            self.ws_cells[i].format.set_top(boxtype)

        for i in [j for j in self.ws_cells if
                j[0] == to_r
                and to_c >= j[1] >= from_c and j[1]
                ]:
            self.ws_cells[i].format.set_bottom(boxtype)

    def write_to_workbook(self):
        logging.debug("Writing content")
        for i in self.ws_cells:
            self.ws_cells[i].write_to_worksheet()


class ExcelCell:
    def __init__(self, row, col, wb, ws):
        self.row = row
        self.col = col
        self.format = wb.add_format({})
        self.content = None
        self.worksheet = ws
        self.is_merged = False
        self.merging_information = None
        self.is_formula = False

    def write_to_worksheet(self):
        if self.is_merged:
            self.worksheet.merge_range(
                self.merging_information["from_r"],
                self.merging_information["from_c"],
                self.merging_information["to_r"],
                self.merging_information["to_c"],
                self.content, self.format)
        elif self.is_formula:
            self.worksheet.write_formula(self.row, self.col, self.content, self.format)
        else:
            self.worksheet.write(self.row, self.col, self.content, self.format)


    