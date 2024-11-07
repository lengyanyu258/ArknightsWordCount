import datetime
import re
from argparse import Namespace
from collections import Counter
from itertools import product
from pathlib import Path
from typing import Any

from .base import Base, Info
from .excel import CellFormatProperties as Props
from .excel import Sheet
from .utils import amend_sheet_list, find_index, find_indices, merge_sheets_list


class Dump(Base):
    __WORDS = "字词数"
    __PUNCTUATION = "标点数"
    __ELLIPSIS = "省略号"
    __COMMANDS = "指令数"
    __split_pattern = re.compile(r"&|\uFF06|/")

    def __init__(
        self,
        config: Namespace,
        args: Namespace,
    ):
        self.__FONT_NAME: str = config.FONT_NAME
        self.__output_file = Path(config.output_file_path)
        self.__name_prefix: list[str] = config.name_prefix
        self.__name_suffix: list[str] = config.name_suffix
        self.__erase_names: list[str] = config.erase_names
        self.__merge_names: list[list[str]] = config.merge_names

        # https://mirrors.tuna.tsinghua.edu.cn/github-release/be5invis/Sarasa-Gothic/LatestRelease/SarasaMonoSlabSC-TTF-Unhinted-1.0.13.7z
        self.__font_path: dict[str, dict[str, str]] = {
            self.__FONT_NAME: {
                "regular": "./tmp/SarasaMonoSlabSC-Regular.ttf",
                "bold": "./tmp/SarasaMonoSlabSC-Bold.ttf",
            }
        }

        today = f"{datetime.date.today():%Y%m%d}"
        self.__xlsx_file = self.__output_file.with_name(
            f"{self.__output_file.stem}_{today}.xlsx"
        )

        self.__debug: bool = args.debug
        self.__style: bool = args.style or args.publish

    def __merge_counter_dict(self, counter: dict[str, dict]):
        """拆分、合并台词，整理台词量统计数据

        Args:
            counter (dict[str, dict]): 台词量统计数据
        """
        # 分离名称
        for origin_name in list(counter.keys()):
            names = re.split(self.__split_pattern, origin_name)
            if len(names) > 1:
                for name in names:
                    if name in counter:
                        for key in counter[name]:
                            counter[name][key] += counter[origin_name][key]
                    else:
                        counter[name] = counter[origin_name]
                del counter[origin_name]

        origin_names = list(counter.keys())
        merged_names: list[list[str]] = []
        self.__name_prefix.append("")
        self.__name_suffix.append("")
        name_prefixes = set(self.__name_prefix)
        name_suffixes = set(self.__name_suffix)

        # 扩增并过滤单个名称
        for person in self.__merge_names:
            names = set()
            for name in person:
                for i in product(name_prefixes, [name], name_suffixes):
                    new_name = "".join(i)
                    if new_name in origin_names:
                        names.add(new_name)
            if len(names) > 1:
                merged_names.append(list(names))
                for new_name in names:
                    origin_names.remove(new_name)

        # 扩增名称
        for origin_name in origin_names[:]:
            # TODO: ❌❌的⭕⭕ -> (❌❌|❌❌)的⭕⭕ or '\n'.join(['❌❌的⭕⭕',...])
            if self.__debug and "的" in origin_name:
                p, s = origin_name.split(sep="的", maxsplit=1)
                print("(suf|pre)fix:", s, p, origin_name)

            names = []
            for i in product(name_prefixes, [origin_name], name_suffixes):
                new_name = "".join(i)
                if new_name in origin_names:
                    names.append(new_name)
            if len(names) > 1:
                merged_names.append(names)
                for new_name in names:
                    origin_names.remove(new_name)

        # 排除名称
        for names in merged_names[:]:
            for name in names[:]:
                if name in self.__erase_names:
                    names.remove(name)
            if len(names) == 1:
                merged_names.remove(names)

        # 合并名称
        counter_dict: dict[str, dict[str, int]] = {}
        for names in merged_names:
            merged_name = "/".join(sorted(names, key=lambda name: len(name.encode())))
            counter_dict[merged_name] = {
                "words": 0,
                "punctuation": 0,
                "ellipsis": 0,
            }
            for name in names:
                for key in counter[name]:
                    counter_dict[merged_name][key] += counter[name][key]
                del counter[name]
            # 过滤掉 Word 与 Punctuation 为 0 的名字
            if sum(counter_dict[merged_name].values()) > 0:
                counter[merged_name] = counter_dict[merged_name]

    def __gen_sorted_counter_data(
        self,
        tab_time: int,
        info_dict: dict[str, dict],
        sheet_list: list,
        number: int | None = 10,
        is_show_counter: bool = True,
        is_show_all: bool = False,
    ):
        """生成按字词数排序的台词量统计

        Args:
            tab_time (int): 左侧留出空白单元格的次数
            info_dict (dict[str, dict]): 统计数据
            sheet_list (list): 需要被添加到的数据表单
            number (int | None, optional): 要生成的总数. Defaults to 10.
            is_show_counter (bool, optional): Print `Counter` cell left beside `Index` title. Defaults to True.
        """

        sheet_list.append(
            [None] * tab_time
            + (["Counter"] if is_show_counter else [])
            + [
                "Index",
                "Name",
                self.__WORDS,
                self.__PUNCTUATION,
                self.__ELLIPSIS,
            ]
        )

        if is_show_all:
            counter = Counter()
            for name in info_dict["counter"]:
                counter.update(
                    {
                        "words": info_dict["counter"][name]["words"],
                        "punctuation": info_dict["counter"][name]["punctuation"],
                        "ellipsis": info_dict["counter"][name]["ellipsis"],
                    }
                )
            info_dict["counter"]["ALL"] = dict(counter)

        sorted_counter_items = sorted(
            info_dict["counter"].items(),
            key=lambda item: item[1]["words"],
            reverse=True,
        )

        for index, item in enumerate(sorted_counter_items[:number]):
            sheet_list.append(
                [None] * tab_time
                + ([None] if is_show_counter else [])
                + [
                    index if is_show_all else index + 1,
                    item[0],
                    item[1]["words"],
                    item[1]["punctuation"],
                    item[1]["ellipsis"],
                ]
            )

    def __add_info_data(self, sheet_list: list):
        def append_list(a, b):
            sheet_list.append([f"{a}：", b])

        sheet_list.append([self.data["info"]["title"]])
        for k, v in self.data["info"]["data"].items():
            if isinstance(v, list):
                append_list(k, v[0])
                for i in v[1:]:
                    sheet_list.append([None, i])
            else:
                append_list(k, v)

        sheet_list.append([])

    def __gen_info_data(
        self,
        tab_time: int,
        info_dict: dict,
        sheet_list: list,
        number: int | None = 10,
        max_number: int | None = None,
    ):
        bar = {
            "name": "Title",
            "type": "Type",
            "words": self.__WORDS,
            "punctuation": self.__PUNCTUATION,
            "ellipsis": self.__ELLIPSIS,
            "commands": self.__COMMANDS,
        }
        # Including Title bar of `counter`.
        bar_count = 1
        for i in bar:
            if info_dict.get(i):
                bar_count += 1
                sheet_list.append([None] * tab_time + [bar[i], info_dict[i]])

        if max_number is not None:
            number = max_number - bar_count

        if len(info_dict["counter"]) != 1:
            self.__gen_sorted_counter_data(tab_time, info_dict, sheet_list, number)

    def __gen_overview_data(
        self, sheet_overview_list: list, dic: dict[str, dict], sorted_info_key: str
    ):
        sheet_overview_list.append(
            [
                "Index",
                "Name",
                self.__WORDS,
                self.__PUNCTUATION,
                self.__ELLIPSIS,
                self.__COMMANDS,
            ]
        )
        items: dict[str, dict[str, dict]] = dic["items"].copy()

        index_offset = 1
        # 添加总计信息
        if "info" in dic:
            index_offset = 0
            items["ALL"] = {"info": dic["info"].copy()}
            items["ALL"]["info"].pop("name", None)

        keys = list(items.keys())
        sorted_keys = sorted(
            keys,
            key=lambda k: items[k]["info"][sorted_info_key],
            reverse=True,
        )
        for index, k in enumerate(sorted_keys):
            info_dict: dict[str, int] = items[k]["info"]
            sheet_overview_list.append(
                [
                    index + index_offset,
                    info_dict["name"] if "name" in info_dict else k,
                    info_dict["words"],
                    info_dict["punctuation"],
                    info_dict["ellipsis"],
                    info_dict["commands"],
                ]
            )

    def __gen_simple_data(
        self, sheet_simple_list: list, dic: dict[str, dict[str, dict[str, dict]]]
    ):
        def append_list(index: str | None, info_dict: dict):
            content_bar = [
                info_dict[key] if key in dic["info"] else None for key in keys_list
            ]
            sheet_simple_list.append([None, index] + content_bar)

        keys_list = ["name", "words", "punctuation", "ellipsis", "commands"]
        title_bar = [
            "Title",
            self.__WORDS,
            self.__PUNCTUATION,
            self.__ELLIPSIS,
            self.__COMMANDS,
        ]
        sheet_simple_list[-1] += [None] + title_bar
        append_list(None, dic["info"])

        for item_key, item in dic["items"].items():
            sheet_simple_list.append([])
            sheet_simple_list.append([item_key] + [None] + title_bar)

            if len(item["items"]) > 1:
                append_list(None, item["info"])

            for k, i in item["items"].items():
                append_list(k, i["info"])

    def __gen_detail_data(
        self,
        sheet_detail_list: list[list[Any]],
        data_dict: dict[str, dict[str, dict[str, dict]]],
    ):
        """生成每个表单上的所有数据"""

        def gen_story(tab_time: int, data: dict[str, dict[str, dict]]):
            info_dict: dict = data["info"]
            if info_dict["words"] + info_dict["punctuation"] <= 0:
                return

            if len(data["items"]) != 1:
                # 2: 既有“行动前”又有“行动后”
                # 0: 已到达最底层，输出该层信息
                self.__gen_info_data(tab_time, info_dict, level_list)
            else:
                # 只包含单个节点（如“幕间”），只输出标题，不与下一层输出重复的 info data，以减轻文档大小
                if "name" in info_dict:
                    # if "name" not in list(data["items"].values())[0]["info"]:
                    level_list.append([None] * tab_time + ["Title", info_dict["name"]])

            for key in data["items"]:
                level_list.append([None] * tab_time + [key])
                gen_story(tab_time + 1, data["items"][key])

        # 生成 info data
        self.__gen_info_data(0, data_dict["info"], sheet_detail_list)

        stories = data_dict["items"]
        for story_name in stories:
            levels_list = []
            for level_name in stories[story_name]["items"]:
                level_list = [[level_name]]
                gen_story(1, stories[story_name]["items"][level_name])
                amend_sheet_list(level_list)
                levels_list.append(level_list)
            story_list = merge_sheets_list(levels_list, False)
            sheet_detail_list.append([])  # 空一行

            if len(stories[story_name]["items"]) == 1:
                # 如果只包含一个关卡，则不生成总摘要信息
                sheet_detail_list.extend([[story_name]] + story_list)
                continue

            story_digest = [[story_name]]
            self.__gen_info_data(
                1,
                stories[story_name]["info"],
                story_digest,
                max_number=len(story_list) - 1,
            )
            amend_sheet_list(story_digest)
            sheet_detail_list.extend(
                merge_sheets_list([story_digest, story_list], False)
            )

    # @Info("gen overview sheet style...")
    def __gen_overview_sheet_style(self, overview: Sheet):
        overview.default_format_properties.update(
            {"font_name": self.__FONT_NAME, "font_size": 14}
        )
        overview.other_props.update({"font_path": self.__font_path})

        for row, row_data in enumerate(overview.cells):
            for idx_column in find_indices(row_data, "Index"):
                overview[row, idx_column].expand().autofit()

                # Name column: Horizontal Alignment Center
                overview[:, idx_column + 1].set_format(Props.center)
                # “省略号”列
                overview[:, idx_column + 4].set_format({"num_format": "(0)"})

                # Index region: Border Line Style
                overview[row, idx_column].expand().set_format(Props.border)

                # Index range: Title Style
                head_range = overview[row, idx_column].expand("right")
                head_range.set_format(Props.title)

                # 带标题的单元格列表
                bold_column = find_index(row_data, self.__WORDS, idx_column + 1)
                if (
                    (title_row := row - 1) >= 0
                    and isinstance(overview.cells[title_row][idx_column], str)
                    and overview.cells[title_row][idx_column + 1] is None
                ):
                    overview[title_row, head_range.slice.col].merge(Props.title)

                    # Title column: Font Bold
                    head_text = overview.cells[title_row][idx_column]
                    if "Merged" in head_text:
                        bold_column = find_index(
                            row_data, self.__COMMANDS, idx_column + 1
                        )

                overview[:, bold_column].set_format(Props.font_bold)

        # First column: Horizontal Alignment Right
        overview[:, 0].set_format(Props.right)

        # Cells 'ALL' region:
        all_region = overview[0, 0].end("down", time=2).current_region
        all_region.autofit()
        all_region[all_region.current.row + 3, 1].set_format({"num_format": "(0)"})
        overview[1:, 0].autofit()

        # Cells 'Title' Range:
        title_range = overview[0, 0].expand()
        title_range[:, -1].set_format(Props.left)
        for row in range(title_range.last_cell.row):
            if "http" in title_range[row + 1, -1].value:
                title_range[row + 1, -1].set_format({"hyperlink": True})
            overview[
                row + 1,
                title_range.last_cell.col : all_region.last_cell.col + 1,
            ].merge(Props.border)

        # Title region: Border Line Style
        title_range.set_format(Props.border)
        overview[0, : all_region.last_cell.col + 1].merge(Props.title)

    @Info("style formatting...")
    def __gen_sheet_style(self, overview: Sheet, simple: Sheet, counter: Sheet):
        # 『概观』表单
        self.__gen_overview_sheet_style(overview=overview)

        # 『总览』表单
        # The default font_size is 11
        simple.default_format_properties.update({"font_name": self.__FONT_NAME})
        simple.other_props.update({"font_path": self.__font_path})
        simple.autofit()

        for row, row_data in enumerate(simple.cells):
            for idx_column in find_indices(row_data, self.__WORDS):
                simple[row, idx_column].expand("right").set_format(Props.center)
                # “省略号”列
                simple[row, idx_column].current_region[row + 1 :, -2].set_format(
                    {"num_format": "(0)"}
                )

                # Left column: Horizontal Alignment Right
                left_cell = simple[row, idx_column].end("left", time=2).current
                simple[:, left_cell.col].set_format(Props.right)

                # Title column: Horizontal Alignment Center
                simple[:, idx_column - 1].set_format(Props.center)

        # 『台词』表单
        counter.default_format_properties.update(
            {"font_name": self.__FONT_NAME, "font_size": 14}
        )
        counter.other_props.update({"font_path": self.__font_path})
        counter.autofit()

        idx_row = 1
        idx_col = find_index(counter.cells[idx_row], "Index")
        # Name column: Horizontal Alignment Center
        counter[:, idx_col + 1].set_format(Props.center)
        # “省略号”列
        counter[:, idx_col + 4].set_format({"num_format": "(0)"})
        # Index region: Border Line Style
        counter[idx_row, idx_col].expand().set_format(Props.border)
        # Title column: Font Bold
        bold_column = find_index(counter.cells[idx_row], self.__WORDS, idx_col + 1)
        counter[:, bold_column].set_format(Props.font_bold)
        # Index range: Title Style
        head_range = counter[idx_row, idx_col].expand("right")
        head_range.set_format(Props.title)
        counter[idx_row - 1, head_range.slice.col].merge(Props.title)

    @Info("Writing to excel...")
    def __write_excel_data(
        self,
        sheets_overview_list: list,
        sheets_simple_list: list,
        sheet_counter_list: list,
        sheets_detail_dict: dict,
    ):
        from xlsxwriter import Workbook

        with Workbook(self.__xlsx_file) as workbook:
            ## 设置 Workbook 文档属性
            # workbook.read_only_recommended()
            workbook.set_properties(
                {
                    "title": self.data["info"]["title"],
                    "author": "; ".join(
                        [
                            re.sub("( ?<.+>)", "", author)
                            for author in self.data["info"]["authors"]
                        ]
                    ),
                    "comments": "Created with Python and XlsxWriter",
                }
            )
            for k, v in self.data["info"]["data"].items():
                if isinstance(v, list):
                    workbook.set_custom_property(k, "; ".join(v))
                else:
                    workbook.set_custom_property(k, v)

            ## 新增表单
            overview = Sheet(
                workbook=workbook,
                name="概观",
                data=merge_sheets_list(sheets_overview_list),
            )
            simple = Sheet(
                workbook=workbook,
                name="总览",
                data=merge_sheets_list(sheets_simple_list),
            )
            counter = Sheet(workbook=workbook, name="台词", data=sheet_counter_list)

            if self.__style:
                self.__gen_sheet_style(overview, simple, counter)

            # 将数据写回到表单
            overview.write()
            simple.write()
            counter.write()
            for key in sheets_detail_dict:
                Sheet(workbook=workbook, name=key, data=sheets_detail_dict[key]).write()

            overview.worksheet.activate()

    @Info("Generating data...")
    def __gen_excel_data(
        self,
        sheets_overview_list: list,
        sheets_simple_list: list,
        sheets_detail_dict: dict,
        sheet_counter_list: list,
    ):
        storys_overview_dict = {"items": {}}

        for entry_type, item_dict in self.data["count"]["items"].items():
            # 『概观』表单
            sheet_overview_list = [[entry_type]]
            self.__gen_overview_data(sheet_overview_list, item_dict, "words")
            amend_sheet_list(sheet_overview_list)
            sheets_overview_list.append(sheet_overview_list)

            # 『总览』表单
            sheet_simple_list = [[entry_type]]
            self.__gen_simple_data(sheet_simple_list, item_dict)
            amend_sheet_list(sheet_simple_list)
            sheets_simple_list.append(sheet_simple_list)

            sheet_detail_list = [[entry_type]]
            self.__gen_detail_data(sheet_detail_list, item_dict)
            amend_sheet_list(sheet_detail_list)
            sheets_detail_dict[entry_type] = sheet_detail_list

            for story_key, story_dict in item_dict["items"].items():
                if story_key in storys_overview_dict["items"]:
                    info_dict = storys_overview_dict["items"][story_key]["info"]
                    for key in ["words", "punctuation", "commands"]:
                        info_dict[key] += story_dict["info"][key]
                else:
                    storys_overview_dict["items"][story_key] = story_dict

        sheet_overview_list = [["Merged"]]
        self.__gen_overview_data(sheet_overview_list, storys_overview_dict, "commands")
        amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)

        # 『台词』表单
        sheet_counter_list.append(["台词量统计"])
        self.__gen_sorted_counter_data(
            0, self.data["count"]["info"], sheet_counter_list, None, False, True
        )
        amend_sheet_list(sheet_counter_list)

    def dump_excel(self) -> Path:
        sheets_overview_list = []
        sheets_simple_list = []
        sheet_counter_list = []
        sheets_detail_dict = {}

        # 初始化台词量统计数据
        self.__merge_counter_dict(self.data["count"]["info"]["counter"])

        sheet_overview_list = []
        # 添加文档首部信息
        self.__add_info_data(sheet_overview_list)
        # 添加总量统计信息
        sheet_overview_list.append(["ALL"])
        self.__gen_info_data(0, self.data["count"]["info"], sheet_overview_list, 13)
        amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)

        # 生成 Excel 表单数据
        self.__gen_excel_data(
            sheets_overview_list,
            sheets_simple_list,
            sheets_detail_dict,
            sheet_counter_list,
        )

        # 写入 Excel 表单数据
        self.__write_excel_data(
            sheets_overview_list,
            sheets_simple_list,
            sheet_counter_list,
            sheets_detail_dict,
        )

        return self.__xlsx_file
