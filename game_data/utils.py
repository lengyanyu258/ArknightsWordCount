from typing import Any, Optional, SupportsIndex

from PIL import ImageFont


def find_index(
    list_to_check: list,
    item_to_find: Any,
    start_col_num: int,
) -> int:
    for index, value in enumerate(list_to_check[start_col_num:], start_col_num):
        if value == item_to_find:
            return index
    return -1


def find_indices(list_to_check: list, item_to_find: Any) -> list[int]:
    return [idx for idx, value in enumerate(list_to_check) if value == item_to_find]


def set_column_cell_format(
    cells_format: list[list[Optional[dict[str, Any]]]],
    column_num: SupportsIndex,
    cell_format: dict[str, Any],
):
    for row_num in range(len(cells_format)):
        cell_format_props = cells_format[row_num][column_num]
        if cell_format_props is None:
            cells_format[row_num][column_num] = cell_format
        else:
            cell_format_props.update(cell_format)


def set_region_cells_format(
    cells_format: list[list[Optional[dict[str, Any]]]],
    first_row_num: int,
    first_col_num: int,
    last_row_num: int,
    last_col_num: int,
    cell_format: dict[str, Any],
):
    for row_num in range(first_row_num, last_row_num + 1):
        for col_num in range(first_col_num, last_col_num + 1):
            cell_format_props = cells_format[row_num][col_num]
            if cell_format_props is not None:
                cell_format_props.update(cell_format)


def get_end_cell_number(
    cells: list[list[Any]],
    row_num: int,
    col_num: int,
    direction: str,
) -> int:
    if direction == "right":
        for column_number in range(col_num + 1, len(cells[row_num])):
            if cells[row_num][column_number] is None:
                return column_number - 1
        else:
            return len(cells[row_num]) - 1
    elif direction == "down":
        for row_number in range(row_num + 1, len(cells)):
            if cells[row_number][col_num] is None:
                return row_number - 1
        else:
            return len(cells) - 1
    else:
        raise ValueError(f"direction: '{direction}' not valid!")


def get_skip_next_cell_number(
    cells: list[list[Any]],
    row_num: int,
    col_num: int,
    direction: str,
) -> int:
    if direction == "left":
        column_range = iter(range(col_num - 1, -1, -1))
        for column_number in column_range:
            if cells[row_num][column_number] is None:
                last_column = column_number + 1
                # Skip 'None' column number
                for column_number in column_range:
                    if cells[row_num][column_number] is not None:
                        return column_number
                else:
                    return last_column
        else:
            return 0
    elif direction == "down":
        row_range = iter(range(row_num + 1, len(cells)))
        for row_number in row_range:
            if cells[row_number][col_num] is None:
                last_row = row_number - 1
                # Skip 'None' row number
                for row_number in row_range:
                    if cells[row_number][col_num] is not None:
                        return row_number
                else:
                    return last_row
        else:
            return len(cells) - 1
    else:
        raise ValueError(f"direction: '{direction}' not valid!")


def get_column_width(
    cells: list[list[Any]],
    column_num: int,
    font_path: str,
    font_size: int,
    start_row_num: int = 0,
) -> float:
    # column width in character units.

    text_list = [
        str(cells[row_num][column_num])
        for row_num in range(start_row_num, len(cells))
        if cells[row_num][column_num] is not None
    ]

    length = 0
    if len(text_list):
        font = ImageFont.truetype(font_path, font_size)
        for text in text_list:
            length = max(font.getlength(text), length)
    else:
        length = 80 * 64 / 125

    return length / 5
