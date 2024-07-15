from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from logging import warn
from typing import Any, Optional, SupportsIndex

from PIL import ImageFont
from xlsxwriter import Workbook
from xlsxwriter.format import Format

from .utils import Axis, check_index


class ExpandType(StrEnum):
    DOWN = "down"
    RIGHT = "right"
    TABLE = "table"


class DirectionType(StrEnum):
    UP = "up"
    DOWN = "down"
    LEFT = "left"
    RIGHT = "right"


@dataclass(init=False, frozen=True)
class CellFormatProperties:
    """单元格格式属性
    Check at https://xlsxwriter.readthedocs.io/format.html
    """

    border = {"border": 1}
    center = {"align": "center", "valign": "vcenter"}
    font_bold = {"bold": True}
    right = {"align": "right", "valign": "vcenter"}
    title = {**center, "border": 2}


@dataclass(frozen=True)
class Index:
    row: int
    col: int


@dataclass(frozen=True)
class RangeIndex:
    first: Index
    last: Index


@dataclass(frozen=True)
class RangeSlice:
    row: slice
    col: slice


def parse_index_tuple(
    index_tuple: tuple[int | slice, int | slice], index_range: RangeIndex
) -> RangeIndex:
    def parse_index(ind: int | slice, rng: Axis) -> Axis:
        if isinstance(ind, int):
            return Axis(check_index(ind, rng))

        if ind.step is not None and ind.step != 1:
            raise NotImplementedError("Not support yet.")

        left, right = rng
        if isinstance(ind.start, int):
            left = check_index(ind.start, rng)
        if isinstance(ind.stop, int):
            # ind(slice): 左闭右开区间 `[start, stop)`
            right = check_index(ind.stop - 1, rng)
            # 因此语义限制，须保证有 left <= right

        return Axis(left, right)

    row = parse_index(index_tuple[0], Axis(index_range.first.row, index_range.last.row))
    col = parse_index(index_tuple[1], Axis(index_range.first.col, index_range.last.col))

    return RangeIndex(Index(row.left, col.left), Index(row.right, col.right))


class Range:
    """Rectangle range"""

    def __init__(
        self,
        sheet: Sheet,
        index: Optional[RangeIndex] = None,
    ):
        self.sheet = sheet

        self.__breadth = len(self.sheet.cells[0])
        self.__depth = len(self.sheet.cells)
        # 未作合法性检验
        self.__index = index or RangeIndex(
            first=Index(row=0, col=0),
            last=Index(row=self.__depth - 1, col=self.__breadth - 1),
        )
        self.__range = RangeSlice(
            row=slice(self.__index.first.row, self.__index.last.row + 1),
            col=slice(self.__index.first.col, self.__index.last.col + 1),
        )
        self.__active_cell = self.__index.first

    def __getitem__(self, index_tuple: tuple[int | slice, int | slice]) -> Range:
        index_range = (
            self.sheet.__index
            if self.__index.first == self.__index.last
            else self.__index
        )
        return Range(
            sheet=self.sheet,
            index=parse_index_tuple(index_tuple, index_range),
        )

    @property
    def slice(self) -> RangeSlice:
        return self.__range

    @property
    def last_cell(self) -> Index:
        return self.__index.last

    @property
    def value(self) -> Any:
        return self.sheet.cells[self.current.row][self.current.col]

    @property
    def current(self) -> Index:
        return self.__active_cell

    @property
    def current_region(self) -> Range:
        """Same as `Ctrl + *` in Excel."""
        first_cell = self.end(DirectionType.LEFT).end(DirectionType.UP)
        while True:
            tried_cell = first_cell.end(DirectionType.LEFT).end(DirectionType.UP)
            if tried_cell.current == first_cell.current:
                break
            first_cell = tried_cell

        last_cell = self.end(DirectionType.DOWN).end(DirectionType.RIGHT)
        while True:
            tried_cell = last_cell.end(DirectionType.DOWN).end(DirectionType.RIGHT)
            if tried_cell.current == last_cell.current:
                break
            last_cell = tried_cell

        return self[
            first_cell.current.row : last_cell.current.row + 1,
            first_cell.current.col : last_cell.current.col + 1,
        ]

    def expand(self, mode: str | ExpandType = ExpandType.TABLE) -> Range:
        match ExpandType(mode):
            case ExpandType.DOWN:
                last_cell = self.end(DirectionType.DOWN)
            case ExpandType.RIGHT:
                last_cell = self.end(DirectionType.RIGHT)
            case ExpandType.TABLE:
                last_cell = self.end(DirectionType.DOWN).end(DirectionType.RIGHT)
            case _:
                raise ValueError(
                    f"Expand mode: `{mode}` is not valid! Should be one of: {ExpandType._member_names_}."
                )
        return self[
            self.current.row : last_cell.current.row + 1,
            self.current.col : last_cell.current.col + 1,
        ]

    def __end_cell_row_number(
        self,
        range_object: range,
        column: SupportsIndex,
        action_time: int,
    ) -> int:
        def goto_end(last_row: int) -> int:
            for row in row_range:
                if self.sheet.cells[row][column] is None:
                    break
                last_row = row
            return last_row

        def skip_blank(last_row: int) -> int:
            for row in row_range:
                if self.sheet.cells[row][column] is not None:
                    return row
            return last_row

        row_range = iter(range_object)
        next_row = next(row_range)

        is_start_blank = self.sheet.cells[next_row][column] is None
        for _ in range(action_time):
            last_row = next_row
            match (_ + is_start_blank) % 2:
                case 0:
                    next_row = goto_end(last_row)
                case 1:
                    next_row = skip_blank(last_row)

        return next_row

    def __end_cell_column_number(
        self,
        range_object: range,
        row: SupportsIndex,
        action_time: int,
    ) -> int:
        def goto_end(last_column: int) -> int:
            for column in column_range:
                if self.sheet.cells[row][column] is None:
                    break
                last_column = column
            return last_column

        def skip_blank(last_column: int) -> int:
            for column in column_range:
                if self.sheet.cells[row][column] is not None:
                    return column
            return last_column

        column_range = iter(range_object)
        next_column = next(column_range)

        is_start_blank = self.sheet.cells[row][next_column] is None
        for _ in range(action_time):
            last_column = next_column
            match (_ + is_start_blank) % 2:
                case 0:
                    next_column = goto_end(last_column)
                case 1:
                    next_column = skip_blank(last_column)

            if last_column == next_column:
                break

        return next_column

    def end(self, direction: str | DirectionType, time: int = 1) -> Range:
        """Same as `Ctrl + Arrow Keys` in Excel."""
        match DirectionType(direction):
            case DirectionType.UP:
                return self[
                    self.__end_cell_row_number(
                        range_object=range(self.current.row, -1, -1),
                        column=self.current.col,
                        action_time=time,
                    ),
                    self.current.col,
                ]
            case DirectionType.DOWN:
                return self[
                    self.__end_cell_row_number(
                        range_object=range(self.current.row, self.__depth),
                        column=self.current.col,
                        action_time=time,
                    ),
                    self.current.col,
                ]
            case DirectionType.LEFT:
                return self[
                    self.current.row,
                    self.__end_cell_column_number(
                        range_object=range(self.current.col, -1, -1),
                        row=self.current.row,
                        action_time=time,
                    ),
                ]
            case DirectionType.RIGHT:
                return self[
                    self.current.row,
                    self.__end_cell_column_number(
                        range_object=range(self.current.col, self.__breadth),
                        row=self.current.row,
                        action_time=time,
                    ),
                ]
            case _:
                raise ValueError(
                    f"Direction: `{direction}` is not valid! Should be one of: {DirectionType._member_names_}."
                )

    def set_format(self, format: dict[str, Any]):
        for row_num in range(*self.slice.row.indices(self.__depth)):
            for col_num in range(*self.slice.col.indices(self.__breadth)):
                self.sheet.props[row_num][col_num].update(format)

    def merge(self, cell_format: dict[str, Any]):
        range_format = self.sheet.props[self.__index.first.row][self.__index.first.col]
        range_format.update(cell_format)
        self.sheet.worksheet.merge_range(
            self.__index.first.row,
            self.__index.first.col,
            self.__index.last.row,
            self.__index.last.col,
            "",
            self.sheet.workbook.add_format(range_format),
        )

    def __get_column_width(
        self,
        text_list,
        column_num: int,
        font_dict: dict[str, dict[str, dict[int, ImageFont.FreeTypeFont]]],
        font_name,
        font_size,
    ) -> float:
        """column width but in character units number."""

        pixels_width = 0.0
        for row_num, text in text_list:
            if self.sheet.props[row_num][column_num].get("bold", False):
                font = font_dict[font_name]["bold"][font_size]
            else:
                font = font_dict[font_name]["regular"][font_size]

            # We can get the approximate correct pixels number.
            pixels_width = max(font.getlength(text), pixels_width)

        # Assume per character unit width is 5 pixels (for longer width of this font used in Cell).
        # (Because the number will be showed in Scientific Notation due to the short width of Cell.)
        return pixels_width / 5 + 1 * 2

    def __get_row_height(
        self,
        text_list,
        row_num: int,
        font_dict: dict[str, dict[str, dict[int, ImageFont.FreeTypeFont]]],
        font_name,
        font_size: int,
    ) -> float:
        """row height pixels with offsetting."""

        pixels_height = 0.0
        for column_num, text in text_list:
            if self.sheet.props[row_num][column_num].get("bold", False):
                font = font_dict[font_name]["bold"][font_size]
            else:
                font = font_dict[font_name]["regular"][font_size]

            bounding_box = dict(
                zip(["left", "top", "right", "bottom"], font.getbbox(text=text))
            )
            pixels_height = max(bounding_box["bottom"], pixels_height)

        return pixels_height + 1

    def autofit(self):
        # TODO: 优化在 WSL1 下的性能表现（行、列排版占用 40s）
        font_name: str = self.sheet.default_format_properties.get(
            "font_name", Format().font_name
        )
        font_size = self.sheet.default_format_properties.get(
            "font_size", Format().font_size
        )

        if not self.sheet.other_props.setdefault("font_init", False):
            if (
                "font_path" not in self.sheet.other_props
                or font_name not in self.sheet.other_props["font_path"]
            ):
                warn("Failed: Not found font path!")
                return

            font_path: dict[str, str] = self.sheet.other_props["font_path"][font_name]

            self.sheet.other_props["font_dict"] = {
                font_name: {
                    "regular": {
                        font_size: ImageFont.truetype(font_path["regular"], font_size)
                    },
                    "bold": {
                        font_size: ImageFont.truetype(font_path["bold"], font_size)
                    },
                }
            }

            self.sheet.other_props["font_init"] = True

        font_dict = self.sheet.other_props["font_dict"]

        for col_num in range(*self.slice.col.indices(self.__breadth)):
            text_list = [
                (row_num, str(self.sheet.cells[row_num][col_num]))
                for row_num in range(*self.slice.row.indices(self.__depth))
                if self.sheet.cells[row_num][col_num] is not None
            ]

            if len(text_list) == 0:
                continue

            column_width = self.__get_column_width(
                text_list=text_list,
                column_num=col_num,
                font_dict=font_dict,
                font_name=font_name,
                font_size=font_size,
            )
            self.sheet.worksheet.set_column(col_num, col_num, column_width)

        for row_num in range(*self.slice.row.indices(self.__depth)):
            text_list = [
                (column_num, str(self.sheet.cells[row_num][column_num]))
                for column_num in range(*self.slice.col.indices(self.__breadth))
                if self.sheet.cells[row_num][column_num] is not None
            ]

            if len(text_list) == 0:
                continue

            row_height = self.__get_row_height(
                text_list, row_num, font_dict, font_name, font_size
            )
            self.sheet.worksheet.set_row(row_num, row_height)


class Sheet(Range):
    def __init__(
        self,
        workbook: Workbook,
        name: str,
        data: list[list[Any]],
        default_format_props: Optional[dict[str, Any]] = None,
        other_props: Optional[dict[str, Any]] = None,
    ):
        self.workbook = workbook
        self.worksheet: Workbook.worksheet_class = self.workbook.add_worksheet(name)

        # 每个单元格对应的数据
        self.cells = data
        # 每个单元格相应的格式
        self.props: list[list[dict[str, Any]]] = [
            [{} for _ in row] for row in self.cells
        ]

        self.default_format_properties = default_format_props or {}
        self.other_props = other_props or {}

        super().__init__(sheet=self)

    def write(self):
        for row, row_data in enumerate(self.cells):
            for col, data in enumerate(row_data):
                if data is None:
                    continue

                format_props = {
                    **self.default_format_properties,
                    **self.props[row][col],
                }
                if len(format_props) == 0:
                    self.worksheet.write(row, col, data)
                else:
                    self.worksheet.write(
                        row,
                        col,
                        data,
                        self.workbook.add_format(format_props),
                    )
