import re
import math
import datetime
from pathlib import Path
from argparse import Namespace
from itertools import product

from .data import Data


class XlBordersIndex:
    # 从区域中每个单元格的左上角到右下角的边框。
    xlDiagonalDown = 5
    # 从区域中每个单元格的左下角到右上角的边框。
    xlDiagonalUp = 6
    # 区域底部的边框。
    xlEdgeBottom = 9
    # 区域左边缘的边框。
    xlEdgeLeft = 7
    # 区域右边缘的边框。
    xlEdgeRight = 10
    # 区域顶部的边框。
    xlEdgeTop = 8
    # 区域中所有单元格的水平边框（区域以外的边框除外）。
    xlInsideHorizontal = 12
    # 区域中所有单元格的垂直边框（区域以外的边框除外）。
    xlInsideVertical = 11


class XlLineStyle:
    # 实线。
    xlContinuous = 1
    # 虚线。
    xlDash = -4115
    # 点划相间线。
    xlDashDot = 4
    # 划线后跟两个点。
    xlDashDotDot = 5
    # 点线。
    xlDot = -4142
    # 双线。
    xlDouble = -4119
    # 无线。
    xlLineStyleNone = -4118
    # 倾斜的划线。
    xlSlantDashDot = 13


class XlBorderWeight:
    # 细线（最细的边框）。
    xlHairline = 1
    # 中。
    xlMedium = -4138
    # 粗（最宽的边框）。
    xlThick = 4
    # 薄。
    xlThin = 2


class XlHAlign:
    # 居中。
    xlHAlignCenter = -4108
    # 跨列居中。
    xlHAlignCenterAcrossSelection = 7
    # 分散对齐。
    xlHAlignDistributed = -4117
    # 填充。
    xlHAlignFill = 5
    # 按数据类型对齐。
    xlHAlignGeneral = 1
    # 两端对齐。
    xlHAlignJustify = -4130
    # 靠左。
    xlHAlignLeft = -4131
    # 靠右。
    xlHAlignRight = -4152


class Dump(Data):
    __WORDS = "字词数"
    __PUNCTUATION = "标点数"
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

        today = f"{datetime.date.today():%Y%m%d}"
        file_name = f"{self.__output_file.stem}_{today}"
        self.__xlsx_file = self.__output_file.with_name(f"{file_name}.xlsx")
        self.__csv_file = self.__output_file.with_name(f"{file_name}.csv")

        self.__debug: bool = args.debug
        self.__style: bool = args.style or args.publish
        self.__show_counter: bool = args.show_counter
        self.__show_total: bool = args.show_total

    def __find_indices(self, list_to_check: list, item_to_find: str | int) -> list:
        return [idx for idx, value in enumerate(list_to_check) if value == item_to_find]

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
            # TODO: ❌❌的⭕⭕ -> (❌❌|❌❌)的⭕⭕
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
                content_bar = [""] + bar
                try:
                    sheet_list[index] += content_bar
                except IndexError:
                    sheet_list.append([""] * len_sheet_bar + content_bar)
            else:
                len_amend_sheet = len(sheet_list[0]) - len_sheet_bar
                content_bar = [""] * len_amend_sheet
                for bar in sheet_list[len(sheet) :]:
                    bar += content_bar
        return sheet_list

    def __amend_sheet_list(self, sheet_list: list[list]):
        len_counter = [len(i) for i in sheet_list]
        maximum_offset = max(len_counter)

        for i in range(len(sheet_list)):
            sheet_list[i] += [""] * (maximum_offset - len_counter[i])

    def __gen_sorted_counter_data(
        self,
        tab_time: int,
        info_dict: dict[str, dict],
        sheet_list: list,
        number: int | None = 10,
        is_show_counter: bool = True,
    ):
        sheet_list.append(
            [""] * tab_time
            + (["Counter"] if is_show_counter else [])
            + [
                "Index",
                "Name",
                self.__WORDS,
                self.__PUNCTUATION,
            ]
        )
        sorted_counter_items = sorted(
            info_dict["counter"].items(),
            key=lambda item: item[1]["words"],
            reverse=True,
        )
        for index, item in enumerate(sorted_counter_items[:number]):
            sheet_list.append(
                [""] * tab_time
                + ([""] if is_show_counter else [])
                + [
                    index + 1,
                    item[0],
                    item[1]["words"],
                    item[1]["punctuation"],
                ]
            )

    def __add_info_data(self, info: dict, sheet_list: list):
        def append_list(a, b):
            sheet_list.append([f"{a}：", f"'{b}"])

        sheet_list.append([info["title"]])
        for k, v in info["data"].items():
            if type(v) == list:
                append_list(k, v[0])
                for i in v[1:]:
                    sheet_list.append(["", f"'{i}"])
            else:
                append_list(k, v)

        sheet_list.append([])

    def __gen_info_data(self, tab_time: int, info_dict: dict, sheet_list: list):
        title_bar = ["name", "type", "words", "punctuation", "commands"]
        for i in title_bar:
            if i in info_dict and info_dict[i]:
                sheet_list.append(
                    [""] * tab_time
                    + [
                        i.title(),
                        info_dict[i],
                    ]
                )

        if len(info_dict["counter"]) != 1:
            self.__gen_sorted_counter_data(tab_time, info_dict, sheet_list)

    def __gen_overview_data(
        self, sheet_overview_list: list, dic: dict[str, dict], sorted_info_key: str
    ):
        sheet_overview_list.append(
            ["Index", "Name", self.__WORDS, self.__PUNCTUATION, "Commands"]
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
                    str(index + 1),
                    info_dict["name"] if "name" in info_dict else k,
                    info_dict["words"],
                    info_dict["punctuation"],
                    info_dict["commands"],
                ]
            )

    def __gen_simple_data(
        self, sheet_simple_list: list, dic: dict[str, dict[str, dict[str, dict]]]
    ):
        def append_list(index: str, info_dict: dict):
            content_bar = [
                info_dict[key] if key in dic["info"] else "" for key in keys_list
            ]
            sheet_simple_list.append(
                [""] + [f"'{index}"] + [f"'{content_bar[0]}"] + content_bar[1:]
            )

        keys_list = ["name", "words", "punctuation", "commands"]
        title_bar = ["Name", self.__WORDS, self.__PUNCTUATION, "Commands"]
        sheet_simple_list[-1] += [""] + title_bar
        append_list("", dic["info"])

        for item_key, item in dic["items"].items():
            sheet_simple_list.append([])
            sheet_simple_list.append([f"'{item_key}"] + [""] + title_bar)

            if len(item["items"]) > 1:
                append_list("", item["info"])

            for k, i in item["items"].items():
                append_list(k, i["info"])

    def __gen_detail_data(self, tab_time: int, dic: dict[str, dict]):
        info_dict: dict = dic["info"]
        if info_dict["words"] + info_dict["punctuation"] > 0:
            items_dict = dic["items"]
            if len(items_dict) == 1:
                if "name" in info_dict:
                    # if "name" not in list(items_dict.values())[0]["info"]:
                    self.__sheet_detail_list.append(
                        [""] * tab_time
                        + [
                            "Name",
                            info_dict["name"],
                        ]
                    )
            else:
                self.__gen_info_data(tab_time, info_dict, self.__sheet_detail_list)

            for key in items_dict:
                self.__sheet_detail_list.append([""] * tab_time + [f"'{key}"])
                self.__gen_detail_data(tab_time + 1, items_dict[key])

    def gen_excel(self, info: dict) -> Path:
        # TODO: use xlsxwriter instead
        import xlwings as xw

        sheets_detail_dict = {}
        sheets_simple_list = []

        self.__merge_counter_dict(self.data["count"]["info"]["counter"])
        sheets_overview_list = []
        sheet_overview_list = []
        self.__add_info_data(info, sheet_overview_list)
        sheet_overview_list.append(["ALL"])
        self.__gen_info_data(0, self.data["count"]["info"], sheet_overview_list)
        self.__amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)
        storys_overview_dict = {"items": {}}

        for entry_type, item_dict in self.data["count"]["items"].items():
            self.__sheet_detail_list = [[entry_type]]
            self.__gen_detail_data(0, item_dict)
            self.__amend_sheet_list(self.__sheet_detail_list)
            sheets_detail_dict[f"{entry_type}"] = self.__sheet_detail_list

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
        else:
            sheet_overview_list = [["Merged"]]
            self.__gen_overview_data(
                sheet_overview_list, storys_overview_dict, "commands"
            )
            self.__amend_sheet_list(sheet_overview_list)
            sheets_overview_list.append(sheet_overview_list)

            sheet_overview_list = [["Counter"]]
            self.__gen_sorted_counter_data(
                0, self.data["count"]["info"], sheet_overview_list, None, False
            )
            self.__amend_sheet_list(sheet_overview_list)
            sheets_overview_list.append(sheet_overview_list)

            sheet_simple_list = self.__merge_sheets_list(sheets_simple_list)
            sheet_overview_list = self.__merge_sheets_list(sheets_overview_list)

        with xw.App(visible=False, add_book=False) as app:
            self._info("Writing to excel...")

            book = app.books.add()
            sheet_overview = book.sheets(1)
            sheet_overview.name = "Overview"
            sheet_overview[0, 0].value = sheet_overview_list

            sheet_simple = book.sheets.add("Simple", after=sheet_overview)
            sheet_simple[0, 0].value = sheet_simple_list

            for key in sheets_detail_dict:
                sheet = book.sheets.add(key, after=book.sheets[-1])
                sheet[0, 0].value = sheets_detail_dict[key]

            if self.__style:
                self._info("style formatting...")

                def add_title_border(
                    range: xw.Range, weight: int = XlBorderWeight.xlMedium
                ):
                    range.api.HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range.api.Borders.Weight = weight

                # 设置字体，耗时操作
                for sheet in book.sheets:
                    sheet.cells.font.name = self.__FONT_NAME
                # 先设置字体大小，会在 autofit() 时计算宽度
                sheet_overview.cells.font.size = 14

                for y, l in enumerate(sheet_overview_list):
                    for x in self.__find_indices(l, "Index"):
                        sheet_overview[y, x].expand().autofit()
                        name_range = sheet_overview[:, x + 1]
                        name_range.api.HorizontalAlignment = XlHAlign.xlHAlignCenter

                        end_cell = sheet_overview[y, x].end("right")

                        if type(sheet_overview[y - 1, x].value) == str:
                            region_range = sheet_overview[y, x].current_region
                            region_range.api.Borders.LineStyle = (
                                XlLineStyle.xlContinuous
                            )

                            title_range = sheet_overview[y - 1, x : end_cell.column]
                            title_range.merge()
                            add_title_border(title_range)
                        else:
                            region_range: xw.Range = sheet_overview[y, x].expand()
                            region_range.api.Borders.LineStyle = (
                                XlLineStyle.xlContinuous
                            )

                        if "Merged" in str(sheet_overview[y - 1, x].value):
                            commands_range = sheet_overview[:, x + 4]
                            commands_range.api.Font.Bold = True
                        else:
                            words_range = sheet_overview[:, x + 2]
                            words_range.api.Font.Bold = True

                        title_range = sheet_overview[y, x : end_cell.column]
                        add_title_border(title_range)

                cell_all = sheet_overview[0, 0].end("down").end("down")
                cell_all.current_region.autofit()
                cell_all_right = cell_all.current_region.last_cell

                sheet_overview[:, 0].api.HorizontalAlignment = XlHAlign.xlHAlignRight
                cell_right: xw.Range = sheet_overview[0, 0].expand().last_cell
                for y in range(cell_right.row):
                    value = str(sheet_overview[y, cell_right.column].value)
                    if "http" in value:
                        sheet_overview[y, cell_right.column].add_hyperlink(value)

                    region_range = sheet_overview[
                        y, cell_right.column : cell_all_right.column
                    ]
                    region_range.merge()
                    region_range.api.Borders.LineStyle = XlLineStyle.xlContinuous

                region_range: xw.Range = sheet_overview[0, 0].expand()
                region_range.api.Borders.LineStyle = XlLineStyle.xlContinuous

                title_range = sheet_overview[0, 0 : cell_all_right.column]
                title_range.merge()
                add_title_border(title_range)

                sheet_simple.autofit()
                for y, l in enumerate(sheet_simple_list):
                    for x in self.__find_indices(l, self.__WORDS):
                        rest_title_range = sheet_simple[y, x : x + 3]
                        rest_title_range.api.HorizontalAlignment = (
                            XlHAlign.xlHAlignCenter
                        )

                for i in range(math.ceil(len(sheet_simple_list[0]) / 7)):
                    # Reduce communicate with excel's times
                    sheet_simple[
                        :, i * 7
                    ].api.HorizontalAlignment = XlHAlign.xlHAlignRight
                    name_range = sheet_simple[:, i * 7 + 2]
                    name_range.api.HorizontalAlignment = XlHAlign.xlHAlignCenter

                self._info("done.", end=True)

            sheet_overview.activate()

            book.save(self.__xlsx_file)
            self._info("Done.", end=True)

        return self.__xlsx_file

    def __get_data(self, output_txt: str, dic: dict, tab_time: int) -> str:
        info_keys = list(dic["info"].keys())
        info_keys.remove("counter")
        if dic["info"]["words"] + dic["info"]["punctuation"] > 0:
            if self.__show_counter:
                for k in info_keys:
                    output_txt += f"{',' * tab_time}{k},{dic['info'][k]}\n"
                sorted_counter_items = sorted(
                    dic["info"]["counter"].items(),
                    # key=lambda item: item[1]["words"] + item[1]["punctuation"],
                    key=lambda item: item[1]["words"],
                    reverse=True,
                )
                output_txt += f"{',' * tab_time}Index,Name,Words,Punctuation\n"
                for index, item in enumerate(sorted_counter_items[:10]):
                    output_txt += f"{',' * tab_time}{index + 1},{item[0]},{item[1]['words']},{item[1]['punctuation']}\n"
            else:
                content_bar = ""
                for k in info_keys:
                    content_bar += f"{dic['info'][k]},"
                output_txt += f"{content_bar},\n"
        if len(dic["items"]) > 1:
            if self.__show_counter:
                for key in dic["items"]:
                    txt = self.__get_data("", dic["items"][key], tab_time + 1)
                    if len(txt):
                        output_txt += f'{"," * tab_time}|{key}|\n'
                        output_txt += txt
            else:
                for key in dic["items"]:
                    if "name" not in dic["items"][key]["info"]:
                        continue
                    txt = self.__get_data("", dic["items"][key], tab_time + 1)
                    if len(txt):
                        if tab_time == 0:
                            output_txt += "\n\n"
                        elif tab_time == 1:
                            title_bar = ""
                            for k in info_keys:
                                title_bar += f"{k},"
                            output_txt += f"\n{',' * (tab_time + 2)}{title_bar}\n"
                        output_txt += f"{',' * tab_time}'{key}"
                        output_txt += f"{',' * (3 - tab_time)}{txt}"
        return output_txt

    def gen_csv(self) -> Path:
        output_txt = self.__get_data("", self.data["count"], 0)

        if self.__show_total:
            output_txt += "\n"
            story_dict_list = []
            for key in self.data["count"]["items"]:
                dic: dict[str, dict] = self.data["count"]["items"][key]
                keys = list(dic["items"].keys())
                output_txt += f'"{key}"\n'
                sorted_keys = sorted(
                    keys,
                    # key=lambda k: dic[k]["info"]["words"] + dic[k]["info"]["punctuation"],
                    key=lambda k: dic["items"][k]["info"]["words"],
                    reverse=True,
                )
                output_txt += f",Index,Name,Words,Punctuation,Commands\n"
                for index, k in enumerate(sorted_keys):
                    if "name" in dic["items"][k]["info"]:
                        name = dic["items"][k]["info"]["name"]
                        story_dict_list.append(dic["items"][k])
                    else:
                        name = k
                    output_txt += f",{index + 1},{name},{dic['items'][k]['info']['words']},{dic['items'][k]['info']['punctuation']},{dic['items'][k]['info']['commands']}\n"

            output_txt += "\n"
            output_txt += "Commands\n"
            output_txt += f",Index,Name,Words,Punctuation,Commands\n"
            sorted_story_dict_list = sorted(
                story_dict_list,
                key=lambda dic: dic["info"]["commands"],
                reverse=True,
            )
            for index, dic in enumerate(sorted_story_dict_list):
                output_txt += f",{index + 1},{dic['info']['name']},{dic['info']['words']},{dic['info']['punctuation']},{dic['info']['commands']}\n"

            output_txt += "\n"
            output_txt += "Counter\n"
            output_txt += f",Index,Name,Words,Punctuation\n"
            sorted_counter_items = sorted(
                self.data["count"]["info"]["counter"].items(),
                key=lambda item: item[1]["words"],
                reverse=True,
            )
            for index, item in enumerate(sorted_counter_items):
                output_txt += f",{index + 1},{item[0]},{item[1]['words']},{item[1]['punctuation']}\n"

        self.__csv_file.write_text(output_txt)

        return self.__csv_file
