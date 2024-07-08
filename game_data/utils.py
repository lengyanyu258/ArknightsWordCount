from dataclasses import dataclass
from typing import Any, Optional


@dataclass(init=False, repr=False, frozen=True)
class Axis:
    """数轴范围端点，闭区间 `[left, right]`"""

    left: int
    right: int

    def __init__(self, left: int, right: Optional[int] = None) -> None:
        if right is None:
            right = left

        if left > right:
            raise ValueError(f"Can not left > right in axis: {left} > {right}!")

        object.__setattr__(self, "left", left)
        object.__setattr__(self, "right", right)

    def __iter__(self):
        for index in (self.left, self.right):
            yield index

    def __repr__(self) -> str:
        return f"[{self.left}, {self.right}]"


def check_index(index: int, index_range: Axis) -> int:
    if index < 0:
        return index % (index_range.right + 1)

    if index < index_range.left or index > index_range.right:
        raise IndexError(f"Out of index: {index} not in range {index_range}!")

    return index


def find_index(
    list_to_check: list[Any],
    item_to_find: Any,
    start_col_num: int = 0,
) -> int:
    for index, value in enumerate(list_to_check[start_col_num:], start_col_num):
        if value == item_to_find:
            return index
    raise ValueError(f"Not found item: `{item_to_find}`!")


def find_indices(list_to_check: list, item_to_find: Any) -> list[int]:
    return [idx for idx, value in enumerate(list_to_check) if value == item_to_find]


def amend_sheet_list(sheet_list: list[list[Any]]):
    """将数据表单修正为矩形

    Args:
        sheet_list (list[list]): 要被修正的数据表单
    """
    len_counter = [len(i) for i in sheet_list]
    maximum_offset = max(len_counter)

    for i in range(len(sheet_list)):
        sheet_list[i] += [None] * (maximum_offset - len_counter[i])


def merge_sheets_list(
    sheets: list[list[list[Any]]], add_pad: bool = True
) -> list[list[Any]]:
    """将多个表单数据按顺序依次合并为一个表单数据

    Args:
        sheets (list[list[list]]): 多个表单数据的列表，每个表单的数据都为矩形

    Returns:
        list[list]: 合并后的单个表单数据
    """
    sheet_list = sheets[0]
    for sheet in sheets[1:]:
        len_sheet_bar = len(sheet_list[0])
        for index, bar in enumerate(sheet):
            content_bar = ([None] if add_pad else []) + bar
            try:
                sheet_list[index].extend(content_bar)
            except IndexError:
                sheet_list.append([None] * len_sheet_bar + content_bar)
        content_bar = [None] * (len(sheet[0]) + add_pad)
        for bar in sheet_list[len(sheet) :]:
            bar.extend(content_bar)
    return sheet_list
