import datetime
import re
from argparse import Namespace
from itertools import product
from pathlib import Path

from .data import Data


class Dump(Data):
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

    def __merge_sheets_list(self, sheets: list[list[list]]) -> list[list]:
        sheet_list = sheets[0]
        for sheet in sheets[1:]:
            len_sheet_bar = len(sheet_list[0])
            for index, bar in enumerate(sheet):
                content_bar = [None] + bar
                try:
                    sheet_list[index] += content_bar
                except IndexError:
                    sheet_list.append([None] * len_sheet_bar + content_bar)
            len_amend_sheet = len(sheet_list[0]) - len_sheet_bar
            content_bar = [None] * len_amend_sheet
            for bar in sheet_list[len(sheet) :]:
                bar += content_bar
        return sheet_list

    def __amend_sheet_list(self, sheet_list: list[list]):
        len_counter = [len(i) for i in sheet_list]
        maximum_offset = max(len_counter)

        for i in range(len(sheet_list)):
            sheet_list[i] += [None] * (maximum_offset - len_counter[i])

    def __gen_sorted_counter_data(
        self,
        tab_time: int,
        info_dict: dict[str, dict],
        sheet_list: list,
        number: int | None = 10,
        is_show_counter: bool = True,
    ):
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
                    index + 1,
                    item[0],
                    item[1]["words"],
                    item[1]["punctuation"],
                    item[1]["ellipsis"],
                ]
            )

    def __add_info_data(self, info: dict, sheet_list: list):
        def append_list(a, b):
            sheet_list.append([f"{a}：", b])

        sheet_list.append([info["title"]])
        for k, v in info["data"].items():
            if isinstance(v, list):
                append_list(k, v[0])
                for i in v[1:]:
                    sheet_list.append([None, i])
            else:
                append_list(k, v)

        sheet_list.append([])

    def __gen_info_data(
        self, tab_time: int, info_dict: dict, sheet_list: list, number: int | None = 10
    ):
        bar = {
            "name": "Title",
            "type": "Type",
            "words": self.__WORDS,
            "punctuation": self.__PUNCTUATION,
            "ellipsis": self.__ELLIPSIS,
            "commands": self.__COMMANDS,
        }
        for i in bar:
            if i in info_dict and info_dict[i]:
                sheet_list.append([None] * tab_time + [bar[i], info_dict[i]])

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
        keys = list(dic["items"].keys())
        sorted_keys = sorted(
            keys,
            key=lambda k: dic["items"][k]["info"][sorted_info_key],
            reverse=True,
        )
        for index, k in enumerate(sorted_keys):
            info_dict = dic["items"][k]["info"]
            sheet_overview_list.append(
                [
                    index + 1,
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

    def __gen_detail_data(self, tab_time: int, dic: dict[str, dict]):
        # TODO: use horizontal format
        info_dict: dict = dic["info"]
        if info_dict["words"] + info_dict["punctuation"] > 0:
            items_dict = dic["items"]
            if len(items_dict) == 1:
                if "name" in info_dict:
                    # if "name" not in list(items_dict.values())[0]["info"]:
                    self.__sheet_detail_list.append(
                        [None] * tab_time + ["Title", info_dict["name"]]
                    )
            else:
                self.__gen_info_data(tab_time, info_dict, self.__sheet_detail_list)

            for key in items_dict:
                self.__sheet_detail_list.append([None] * tab_time + [key])
                self.__gen_detail_data(tab_time + 1, items_dict[key])

    def gen_excel(self, info: dict) -> Path:
        from .excel import CellFormatProperties as Props
        from .excel import Sheet, Workbook

        sheets_overview_list = []
        sheets_simple_list = []
        sheets_detail_dict = {}

        self.__merge_counter_dict(self.data["count"]["info"]["counter"])
        sheet_overview_list = []
        self.__add_info_data(info, sheet_overview_list)
        sheet_overview_list.append(["ALL"])
        self.__gen_info_data(0, self.data["count"]["info"], sheet_overview_list, 13)
        self.__amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)
        storys_overview_dict = {"items": {}}

        for entry_type, item_dict in self.data["count"]["items"].items():
            self.__sheet_detail_list = [[entry_type]]
            self.__gen_detail_data(0, item_dict)
            self.__amend_sheet_list(self.__sheet_detail_list)
            sheets_detail_dict[entry_type] = self.__sheet_detail_list

            sheet_simple_list = [[entry_type]]
            self.__gen_simple_data(sheet_simple_list, item_dict)
            self.__amend_sheet_list(sheet_simple_list)
            sheets_simple_list.append(sheet_simple_list)

            sheet_overview_list = [[entry_type]]
            self.__gen_overview_data(sheet_overview_list, item_dict, "words")
            self.__amend_sheet_list(sheet_overview_list)
            sheets_overview_list.append(sheet_overview_list)

            for story_key, story_dict in item_dict["items"].items():
                if story_key in storys_overview_dict["items"]:
                    info_dict = storys_overview_dict["items"][story_key]["info"]
                    for key in ["words", "punctuation", "commands"]:
                        info_dict[key] += story_dict["info"][key]
                else:
                    storys_overview_dict["items"][story_key] = story_dict

        sheet_overview_list = [["Merged"]]
        self.__gen_overview_data(sheet_overview_list, storys_overview_dict, "commands")
        self.__amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)

        sheet_overview_list = [["Counter"]]
        self.__gen_sorted_counter_data(
            0, self.data["count"]["info"], sheet_overview_list, None, False
        )
        self.__amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)

        with Workbook(self.__xlsx_file) as workbook:
            self._info("Writing to excel...")

            ## 设置 Workbook 文档属性
            # workbook.read_only_recommended()
            workbook.set_properties(
                {
                    "title": info["title"],
                    "author": "; ".join(
                        [re.sub("( ?<.+>)", "", author) for author in info["authors"]]
                    ),
                    "comments": "Created with Python and XlsxWriter",
                }
            )
            for k, v in info["data"].items():
                if isinstance(v, list):
                    workbook.set_custom_property(k, "; ".join(v))
                else:
                    workbook.set_custom_property(k, v)

            ## 新增表单
            overview = Sheet(
                workbook=workbook,
                name="概观",
                data=self.__merge_sheets_list(sheets_overview_list),
                default_format_props={
                    "font_name": self.__FONT_NAME,
                    "font_size": 14,
                },
                other_props={"font_path": self.__font_path},
            )
            simple = Sheet(
                workbook=workbook,
                name="总览",
                data=self.__merge_sheets_list(sheets_simple_list),
                default_format_props={
                    "font_name": self.__FONT_NAME,
                    # The default font_size is 11
                    # "font_size": 11,
                },
                other_props={"font_path": self.__font_path},
            )

            if self.__style:
                self._info("style formatting...")

                from .utils import find_index, find_indices

                for row, row_data in enumerate(overview.cells):
                    for idx_column in find_indices(row_data, "Index"):
                        overview[row, idx_column].expand().autofit()

                        # Name column: Horizontal Alignment Center
                        overview[:, idx_column + 1].set_format(Props.center)

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
                overview[1:, 0].autofit()

                # Cells 'Title' Range:
                title_range = overview[0, 0].expand()
                cell_col = title_range.last_cell.col
                for row in range(title_range.last_cell.row):
                    cell_row = row + 1
                    if "http" in overview[cell_row, cell_col].value:
                        overview[cell_row, cell_col].set_format({"hyperlink": True})
                    overview[
                        cell_row,
                        cell_col : all_region.last_cell.col + 1,
                    ].merge(Props.border)

                # Title region: Border Line Style
                title_range.set_format(Props.border)
                overview[0, : all_region.last_cell.col + 1].merge(Props.title)

                overview.write()

                for row, row_data in enumerate(simple.cells):
                    for idx_column in find_indices(row_data, self.__WORDS):
                        simple[row, idx_column].expand("right").set_format(Props.center)

                        # Left column: Horizontal Alignment Right
                        left_cell = simple[row, idx_column].end("left", time=2).current
                        simple[:, left_cell.col].set_format(Props.right)

                        # Title column: Horizontal Alignment Center
                        simple[:, idx_column - 1].set_format(Props.center)

                # Autofit
                simple.autofit()
                simple.write()

                self._info("done.", end=True)
            else:
                overview.write()
                simple.write()

            for key in sheets_detail_dict:
                Sheet(workbook=workbook, name=key, data=sheets_detail_dict[key]).write()

            overview.worksheet.activate()

            self._info("Done.", end=True)

        return self.__xlsx_file
