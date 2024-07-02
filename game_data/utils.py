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
    list_to_check: list,
    item_to_find: Any,
    start_col_num: int,
) -> int:
    for index, value in enumerate(list_to_check[start_col_num:], start_col_num):
        if value == item_to_find:
            return index
    raise ValueError(f"Not found item: `{item_to_find}`!")


def find_indices(list_to_check: list, item_to_find: Any) -> list[int]:
    return [idx for idx, value in enumerate(list_to_check) if value == item_to_find]
