import os
import re
import copy
import math
import json
import pickle
import string
import pathlib
import argparse
import datetime
import platform
import warnings
import zhon.hanzi
import collections
import xlwings as xw
from tqdm import tqdm


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


DATA_PATH = pathlib.Path("./GitHub/ArknightsGameData", "zh_CN/gamedata")

pickle_path = pathlib.Path("tmp", "Arknights_gamedata.pkl")
unknown_files = []
unknown_commands = []
known_commands = (
    "AddItem",
    "background",
    "Background",
    "backgroundtween",
    "backgroundTween",
    "Backgroundtween",
    "BackgroundTween",
    "bgeffect",
    "bgEffect",
    "blocker",
    "Blocker",
    "cameraEffect",
    "CameraEffect",
    "camerashake",
    "CameraShake",
    "chaa",
    "character",
    "Character",
    "characteraction",
    "Characteraction",
    "CharacterCutin",
    "charslot",
    "Charslot",
    "charslsot",
    "Condition",
    "ConsumeGuideOnStoryEnd",
    "curtain",
    "dalay",
    "daley",
    "dealy",
    "delat",
    "delau",
    "delay",
    "Delay",
    "delay9ti",
    "delayt",
    "dialo",
    "effect",
    "Effect",
    "End",
    "fadetime",
    "GotoPage",
    "gridbg",
    "header",
    "hideitem",
    "hideItem",
    "HideItem",
    "image",
    "Image",
    "imagerotate",
    "imageTween",
    "ImageTween",
    "largebg",
    "largebgtween",
    "musicvolume",
    "Musicvolume",
    "MusicVolume",
    "musicvolune",
    "Obtain",
    "OptionBranch",
    "palysound",
    "playmusic",
    "playMusic",
    "PlayMusic",
    "playsound",
    "playSound",
    "PlaySound",
    "predicate",
    "Predicate",
    "SetConditionProgress",
    "showitem",
    "Showitem",
    "ShowItem",
    "skipnode",
    "SkipToThis",
    "soundvolume",
    "soundVolume",
    "SoundVolume",
    "StartBattle",
    "stickerclear",
    "stopmucis",
    "stopmusic",
    "Stopmusic",
    "StopMusic",
    "stopsound",
    "stopSound",
    "Stopsound",
    "StopSound",
    "subtitle",
    "theater",
    "Tutorial",
    "verticalbg",
    "Video",
    "withdraw",
)

text_pattern = re.compile(r"(?:\[(.+)\])?(.+)?", re.MULTILINE)
line_pattern = re.compile(
    r"([AaPp]\.?[Mm]\.?)|([\u03B1-\u03C9\d]+(?:[\.\-:：]\d+)?(?:%|℃|u/L|M)?)|((?:[A-Za-z]+\.(?!\.))+[A-Za-z]*)|([A-Za-z\u03B1-\u03C9ò-ö\d]+(?:[/\-']?[A-Za-z\d]+)?)"
)
command_pattern = re.compile(r"\w+")
subtitle_pattern = re.compile(r"(<[A-Za-z\d/=#@\.]+>)|({@[Nn]ickname})|(\\r)|(\\n)")
split_pattern = re.compile(r"&|\uFF06|/")
punctuation_set = set(string.punctuation + zhon.hanzi.punctuation)
ASIDE_NAME = "“旁白”"

DATA: dict[str, dict[str, dict]] = {
    "excel": {
        "activity_table": {},
        "gamedata_const": {
            "dataVersion": "0.0.0",
        },
        "story_review_table": {},
    },
    "story": {},
    "count": {"info": {}, "items": {}},
}

merge_name_list = [
    ("艾丽妮", "审判官艾丽妮"),
    ("教宗", "伊万杰利斯塔十一世"),
    ("菲亚梅塔", "微光守夜人", "不叫微光守夜人的黎博利"),
    ("陈", "陈晖洁"),
    ("发言人恰尔内", "恰尔内"),
    ("“焰尾”索娜", "索娜"),
    ("灰毫骑士", "灰毫"),
    ("玛嘉烈", "临光"),
    ("玛莉娅", "瑕光"),
    ("祖玛玛", "森蚺"),
    ("发言人马克维茨", "马克维茨"),
    ("塞诺蜜", "砾"),
    ("银灰", "恩希欧迪斯"),
    ("阿黛尔", "艾雅法拉"),
    ("苍苔", "埃尼斯"),
]


def gen_excel():
    def find_indices(list_to_check: list, item_to_find: str | int) -> list:
        return [idx for idx, value in enumerate(list_to_check) if value == item_to_find]

    def merge_counter_dict(counter: dict):
        origin_names = list(counter.keys())
        for origin_name in origin_names:
            names = re.split(split_pattern, origin_name)
            if len(names) > 1:
                for name in names:
                    if name in counter:
                        for key in counter[name]:
                            counter[name][key] += counter[origin_name][key]
                    else:
                        counter[name] = counter[origin_name]
                del counter[origin_name]

        counter_dict: dict[str, dict[str, int]] = {}
        for person in merge_name_list:
            merged_name = "/".join(person)
            counter_dict[merged_name] = {
                "words": 0,
                "punctuation": 0,
            }
            for name in person:
                if name in counter:
                    for key in counter[name]:
                        counter_dict[merged_name][key] += counter[name][key]
                    del counter[name]
            if sum(counter_dict[merged_name].values()) > 0:
                counter[merged_name] = counter_dict[merged_name]

    def merge_sheets_list(sheets: list[list[list]]) -> list[list]:
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

    def amend_sheet_list(sheet_list: list[list]):
        len_counter = [len(i) for i in sheet_list]
        maximum_offset = max(len_counter)

        for i in range(len(sheet_list)):
            sheet_list[i] += [""] * (maximum_offset - len_counter[i])

    def gen_sorted_counter_data(
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
                "Words",
                "Punctuation",
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

    def gen_info_data(tab_time: int, info_dict: dict, sheet_list: list):
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
            gen_sorted_counter_data(tab_time, info_dict, sheet_list)

    def gen_overview_data(dic: dict[str, dict], sorted_info_key: str):
        sheet_overview_list.append(
            ["Index", "Name", "Words", "Punctuation", "Commands"]
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

    def gen_simple_data(dic: dict[str, dict[str, dict[str, dict]]]):
        keys_list = ["name", "words", "punctuation", "commands"]
        title_bar = [key.title() for key in keys_list]
        content_bar = [
            dic["info"][key] if key in dic["info"] else "" for key in keys_list
        ]
        sheet_simple_list[-1] += [""] + title_bar
        sheet_simple_list.append([""] * 2 + content_bar)

        for item_key, item in dic["items"].items():
            sheet_simple_list.append([])
            sheet_simple_list.append([f"'{item_key}"] + [""] + title_bar)

            if len(item["items"]) > 1:
                content_bar = [
                    item["info"][key] if key in dic["info"] else "" for key in keys_list
                ]
                sheet_simple_list.append([""] * 2 + content_bar)

            for k, i in item["items"].items():
                content_bar = [
                    i["info"][key] if key in dic["info"] else "" for key in keys_list
                ]
                sheet_simple_list.append([""] + [f"'{k}"] + content_bar)

    def gen_detail_data(tab_time: int, dic: dict[str, dict]):
        info_dict: dict = dic["info"]
        if info_dict["words"] + info_dict["punctuation"] > 0:
            items_dict = dic["items"]
            if len(items_dict) == 1:
                if "name" in info_dict:
                    # if "name" not in list(items_dict.values())[0]["info"]:
                    sheet_detail_list.append(
                        [""] * tab_time
                        + [
                            "Name",
                            info_dict["name"],
                        ]
                    )
            else:
                gen_info_data(tab_time, info_dict, sheet_detail_list)

            for key in items_dict:
                sheet_detail_list.append([""] * tab_time + [f"'{key}"])
                gen_detail_data(tab_time + 1, items_dict[key])

    sheets_detail_dict = {}
    sheets_simple_list = []

    merge_counter_dict(DATA["count"]["info"]["counter"])
    sheets_overview_list = [[["ALL"]]]
    gen_info_data(0, DATA["count"]["info"], sheets_overview_list[0])
    amend_sheet_list(sheets_overview_list[0])
    storys_overview_dict = {"items": {}}

    for entry_type, item_dict in DATA["count"]["items"].items():
        sheet_detail_list = [[entry_type]]
        gen_detail_data(0, item_dict)
        amend_sheet_list(sheet_detail_list)
        sheets_detail_dict[f"{entry_type}"] = sheet_detail_list

        sheet_simple_list = [[entry_type]]
        gen_simple_data(item_dict)
        amend_sheet_list(sheet_simple_list)
        sheets_simple_list.append(sheet_simple_list)

        sheet_overview_list = [[entry_type]]
        gen_overview_data(item_dict, "words")
        amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)

        for story_key, story_dict in item_dict["items"].items():
            if story_key in storys_overview_dict["items"]:
                info_dict = storys_overview_dict["items"][story_key]["info"]
                for key in ["words", "punctuation", "commands"]:
                    info_dict[key] += story_dict["info"][key]
            else:
                storys_overview_dict["items"][story_key] = story_dict
    else:
        sheet_overview_list = [["Merged Commands"]]
        gen_overview_data(storys_overview_dict, "commands")
        amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)

        sheet_overview_list = [["Counter"]]
        gen_sorted_counter_data(
            0, DATA["count"]["info"], sheet_overview_list, None, False
        )
        amend_sheet_list(sheet_overview_list)
        sheets_overview_list.append(sheet_overview_list)

        sheet_simple_list = merge_sheets_list(sheets_simple_list)
        sheet_overview_list = merge_sheets_list(sheets_overview_list)

    with xw.App(visible=False, add_book=False) as app:
        print("Writing to excel...")

        book = app.books.add()
        sheet_overview = book.sheets(1)
        sheet_overview.name = "Overview"
        sheet_overview[0, 0].value = sheet_overview_list

        sheet_simple = book.sheets.add("Simple", after=sheet_overview)
        sheet_simple[0, 0].value = sheet_simple_list

        for key in sheets_detail_dict:
            sheet = book.sheets.add(key, after=book.sheets[-1])
            sheet[0, 0].value = sheets_detail_dict[key]

        if RESULTS.style:
            # 设置字体，耗时操作
            for sheet in book.sheets:
                sheet.cells.font.name = "Sarasa Mono Slab SC"
            # 先设置字体大小，会在 autofit() 时计算宽度
            sheet_overview.cells.font.size = 14

            for y, l in enumerate(sheet_overview_list):
                for x in find_indices(l, "Index"):
                    sheet_overview[y, x].expand().autofit()
                    name_range = sheet_overview[:, x + 1]
                    name_range.api.HorizontalAlignment = XlHAlign.xlHAlignCenter

                    end_cell = sheet_overview[y, x].end("right")

                    if type(sheet_overview[y - 1, x].value) == str:
                        region_range = sheet_overview[y, x].current_region
                        region_range.api.Borders.LineStyle = XlLineStyle.xlContinuous

                        title_range = sheet_overview[y - 1, x : end_cell.column]
                        title_range.merge()
                        title_range.api.HorizontalAlignment = XlHAlign.xlHAlignCenter
                        title_range.api.Borders.Weight = XlBorderWeight.xlMedium
                    else:
                        region_range = sheet_overview[y, x].expand()
                        region_range.api.Borders.LineStyle = XlLineStyle.xlContinuous

                    if "Commands" in str(sheet_overview[y - 1, x].value):
                        commands_range = sheet_overview[:, x + 4]
                        commands_range.font.name = "Sarasa Mono Slab SC Semibold"
                    else:
                        words_range = sheet_overview[:, x + 2]
                        words_range.font.name = "Sarasa Mono Slab SC Semibold"

                    title_range = sheet_overview[y, x : end_cell.column]
                    title_range.api.HorizontalAlignment = XlHAlign.xlHAlignCenter
                    title_range.api.Borders.Weight = XlBorderWeight.xlMedium

            sheet_overview[0, 0].current_region.autofit()

            sheet_simple.autofit()
            for y, l in enumerate(sheet_simple_list):
                for x in find_indices(l, "Words"):
                    rest_title_range = sheet_simple[y, x : x + 3]
                    rest_title_range.api.HorizontalAlignment = XlHAlign.xlHAlignCenter

            for i in range(math.ceil(len(sheet_simple_list[0]) / 7)):
                # Reduce communicate times
                sheet_simple[:, i * 7].api.HorizontalAlignment = XlHAlign.xlHAlignRight
                name_range = sheet_simple[:, i * 7 + 2]
                name_range.api.HorizontalAlignment = XlHAlign.xlHAlignCenter

        sheet_overview.activate()

        new_name = f"{pickle_path.stem}_count_{datetime.date.today():%Y%m%d}"
        book.save(pickle_path.with_name(f"{new_name}.xlsx"))
        print("done.")


def print_data():
    def get_data(output_txt: str, dic: dict, tab_time: int) -> str:
        info_keys = list(dic["info"].keys())
        info_keys.remove("counter")
        if dic["info"]["words"] + dic["info"]["punctuation"] > 0:
            if RESULTS.show_counter:
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
            if RESULTS.show_counter:
                for key in dic["items"]:
                    txt = get_data("", dic["items"][key], tab_time + 1)
                    if len(txt):
                        output_txt += f'{"," * tab_time}|{key}|\n'
                        output_txt += txt
            else:
                for key in dic["items"]:
                    if "name" not in dic["items"][key]["info"]:
                        continue
                    txt = get_data("", dic["items"][key], tab_time + 1)
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

    output_txt = get_data("", DATA["count"], 0)

    if RESULTS.show_total:
        output_txt += "\n"
        story_dict_list = []
        for key in DATA["count"]["items"]:
            dic: dict[str, dict] = DATA["count"]["items"][key]
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
            DATA["count"]["info"]["counter"].items(),
            key=lambda item: item[1]["words"],
            reverse=True,
        )
        for index, item in enumerate(sorted_counter_items):
            output_txt += (
                f",{index + 1},{item[0]},{item[1]['words']},{item[1]['punctuation']}\n"
            )

    today = f"{datetime.date.today():%Y%m%d}"
    pickle_path.with_name(f"{pickle_path.stem}_{today}.csv").write_text(output_txt)


def count_story(
    command_count: int,
    collection_dict: dict[str, collections.Counter],
    dict_list: list[dict],
):
    words_collection = collections.Counter()
    punctuation_collection = collections.Counter()
    counter_dict: dict[str, dict[str, int]] = {}
    for name, collection in collection_dict.items():
        counter_dict[name] = {
            "words": 0,
            "punctuation": 0,
        }
        collection_set = set(collection)
        for i in punctuation_set & collection_set:
            punctuation_collection.update({i: collection[i]})
            counter_dict[name]["punctuation"] += collection[i]
            del collection[i]
        counter_dict[name]["words"] += collection.total()
        words_collection.update(collection)

    count_dict: dict[str, int] = {
        "commands": command_count,
        "words": words_collection.total(),
        "punctuation": punctuation_collection.total(),
        # "counter": counter_dict,
    }
    for dic in dict_list:
        for key in count_dict:
            if key not in dic:
                dic[key] = count_dict[key]
            else:
                dic[key] += count_dict[key]

        if "counter" not in dic:
            dic["counter"] = copy.deepcopy(counter_dict)
        else:
            for name in counter_dict:
                if name not in dic["counter"]:
                    dic["counter"][name] = counter_dict[name].copy()
                else:
                    for key in counter_dict[name]:
                        dic["counter"][name][key] += counter_dict[name][key]


def parse_line(command: str, text: str) -> tuple[bool, str, collections.Counter]:
    def get_attribute(cmd_str: str):
        # TODO: use regex
        return cmd_str.split(",")[0].split("=")[1].strip(' ")')

    is_command = True
    control_command = command_pattern.match(command)
    if control_command is None:
        control_command = command
    else:
        control_command = control_command.group()

    if control_command in known_commands or command.startswith("[character"):
        return True, "", collections.Counter()
    elif control_command in (
        "HEADER",
        "Title",
        "Div",
    ):
        return False, "", collections.Counter()
    elif control_command in ("Dialog", "PopupDialog", "VoiceWithin", "dialog"):
        # TODO: dialog(head="npc_694_1" 文 activity_table charCardMap
        try:
            head = get_attribute(command)
        except IndexError:
            return True, "", collections.Counter()
        if control_command == "PopupDialog":
            head = DATA["excel"]["story_variables"][head.lstrip("$")]
        if head.startswith("char"):
            try:
                name = (
                    DATA["excel"]["handbook_info_table"]["handbookDict"][head][
                        "storyTextAudio"
                    ][0]["stories"][0]["storyText"]
                    .split("\n")[0]
                    .replace("【代号】", "")
                )
            except KeyError:
                if head in ("char_340_shwazr6"):
                    name = "黑"
                else:
                    warnings.warn(f"not found {head}")
                    name = head
        else:
            name = head
            if head not in unknown_commands:
                unknown_commands.append(head)
    elif control_command in ("name") or command.startswith("(name"):
        is_command = False
        name = get_attribute(command)
        if name == "":
            name = ASIDE_NAME
    elif control_command in ("Decision"):
        # is_command = False
        name = "Dr."
        for i in command.split(","):
            if "option" in i:
                text += "".join(i.split("=")[1].strip(' "').split(";"))
    elif control_command in ("Sticker", "Subtitle"):
        name = ASIDE_NAME
        for i in command.split(","):
            if "text" in i:
                text += i.split("text=")[1].strip(' "')
    elif control_command in ("narration"):
        name = ASIDE_NAME
    elif control_command in ("multiline"):
        name = get_attribute(command)
    else:
        name = ""
        if control_command not in unknown_commands:
            warnings.warn(f"unknwn command: {command}")
            unknown_commands.append(control_command)

    match_set = set()
    for i in subtitle_pattern.finditer(text):
        match_set.add(i.group())
    for i in match_set:
        # print(i, "|", text)
        text = text.replace(i, " ")
    # if len(match_set):
    #     print(text, "\n")

    words = []
    clean_text = ""
    endpos = 0
    for i in line_pattern.finditer(text):
        word = i.group()
        clean_text += text[endpos : i.start()]
        endpos = i.end()
        words.append(word)
    clean_text += text[endpos:]
    # if len(words):
    #     print(words, command)
    #     print(text)
    #     print(clean_text)
    #     print()
    #     temp = (
    #         clean_text.replace(" ", "")
    #         .replace("/", "")
    #         .replace(".", "")
    #         .replace("~", "")
    #         .replace("—", "")
    #         .replace("·", "")
    #         .replace("\\", "")
    #     )
    #     if len(temp.encode("utf-8")) % 3 != 0:
    #         unknown_commands.append(f"\n{words} {command}\n{text}\n{temp}\n")
    collection = collections.Counter(words)
    collection.update(clean_text.replace(" ", ""))
    return is_command, name, collection


def parse_story(story: dict):
    command_count = 0
    collection_dict: dict[str, collections.Counter] = {}

    if RESULTS.count_info:
        txts = story.values()
    else:
        txts = (story["txt"],)

    for txt in txts:
        for line in text_pattern.finditer(txt):
            if line.group() == "":
                continue
            command, text = line.groups()
            if command is None:
                command = f'name="{ASIDE_NAME}"'
            else:
                command = command.strip()
            if text is None:
                text = ""
            else:
                text = text.strip()
            is_command, name, collection = parse_line(command, text)
            if is_command:
                command_count += 1
            if collection.total() > 0:
                if name in collection_dict:
                    collection_dict[name].update(collection)
                else:
                    collection_dict[name] = collection

    return command_count, collection_dict


def count_words():
    print("counting words...")

    DATA["count"] = {"info": {}, "items": {}}
    stories = list(DATA["story"].keys())
    for story_id, story in tqdm(DATA["excel"]["story_review_table"].items()):
        name: str = story["name"]
        entry_type = story["entryType"]
        act_type = story["actType"]
        for infoUnlockData in story["infoUnlockDatas"]:
            story_code: str = infoUnlockData["storyCode"]
            if story_code == "":
                # For mini story
                story_code = str(infoUnlockData["storySort"])
            elif story_code is None:
                # For 人员密录
                story_code = story_id.split("_")[-1]
            story_name: str = infoUnlockData["storyName"]
            story_key: str = infoUnlockData["storyTxt"]
            avg_tag: str = infoUnlockData["avgTag"]
            stories.remove(story_key)
            command_count, collection_dict = parse_story(DATA["story"][story_key])
            if len(collection_dict) == 0:
                continue

            entry_type_dict: dict[str, dict] = DATA["count"]["items"].setdefault(
                entry_type, {"info": {"name": act_type}, "items": {}}
            )
            story_id_dict: dict[str, dict] = entry_type_dict["items"].setdefault(
                story_id, {"info": {"name": name}, "items": {}}
            )
            story_dict: dict[str, dict] = story_id_dict["items"].setdefault(
                story_code, {"info": {"name": story_name}, "items": {}}
            )
            avg_dict: dict[str, dict] = story_dict["items"].setdefault(
                avg_tag, {"info": {}, "items": {}}
            )
            count_story(
                command_count,
                collection_dict,
                [
                    DATA["count"]["info"],
                    entry_type_dict["info"],
                    story_id_dict["info"],
                    story_dict["info"],
                    avg_dict["info"],
                ],
            )

    basicInfo = DATA["excel"]["activity_table"]["basicInfo"]
    banned_dirname = {"guide", "tutorial", "training", "act1bossrush", "bossrush"}
    for story_key in stories:
        parts = story_key.split("/")
        if banned_dirname & set(parts):
            continue
        command_count, collection_dict = parse_story(DATA["story"][story_key])
        if len(collection_dict) == 0:
            continue

        # dic = DATA["count"]["items"].setdefault("OTHERS", {"info": {}, "items": {}})
        dic = DATA["count"]
        dict_list = [dic["info"]]
        for i in parts:
            if i not in dic["items"]:
                if i in basicInfo:
                    info = {
                        "name": basicInfo[i]["name"],
                        "type": basicInfo[i]["type"],
                    }
                else:
                    info = {}
                dic["items"][i] = {"info": info, "items": {}}
            dict_list.append(dic["items"][i]["info"])
            dic = dic["items"][i]
        count_story(command_count, collection_dict, dict_list)

    if len(unknown_commands):
        tmp_text = ""
        for i in unknown_commands:
            tmp_text += f'"{i}",\n'
        pickle_path.with_suffix(".unknown_commands.txt").write_text(
            tmp_text, encoding="utf-8"
        )

    print("done.")


def update_story():
    print("updating story...")

    info_files = STORY_DATA_PATHS["info"].rglob("*.txt")
    activity_files = STORY_DATA_PATHS["activities"].rglob("*.txt")
    obt_files = STORY_DATA_PATHS["obt"].rglob("*.txt")
    files = list(activity_files) + list(obt_files)
    DATA["story"] = {}
    for info in info_files:
        info_relative_path = info.relative_to(STORY_DATA_PATHS["info"])
        story_key = info_relative_path.with_suffix("").as_posix()
        file = STORY_DATA_PATH / info_relative_path
        if file in files:
            files.remove(file)
            txt = file.read_text(encoding="utf-8")
        else:
            warnings.warn(f"{file} not found!")
            unknown_files.append(file.as_posix())
            txt = ""
        DATA["story"][story_key] = {
            "info": info.read_text(encoding="utf-8"),
            "txt": txt,
        }
    else:
        for file in files:
            file_relative_path = file.relative_to(STORY_DATA_PATH)
            story_key = file_relative_path.with_suffix("").as_posix()
            DATA["story"][story_key] = {
                "info": "",
                "txt": file.read_text(encoding="utf-8"),
            }

    if len(unknown_files):
        tmp_text = ""
        for i in unknown_files:
            tmp_text += f'"{i}",\n'
        pickle_path.with_suffix(".unknown_files.txt").write_text(tmp_text)

    print("done.")


def update_data():
    print("updating...")

    for i in EXCEL_DATA_PATHS:
        DATA["excel"][i] = json.loads(EXCEL_DATA_PATHS[i].read_bytes())

    update_story()
    RESULTS.update = False
    count_words()
    RESULTS.count = False

    pickle_path.write_bytes(pickle.dumps(DATA))

    print("done.")


def load_data():
    print("loading...")

    global DATA

    if pickle_path.exists():
        DATA = pickle.loads(pickle_path.read_bytes())

    data_version = DATA_VERSION_PATH.read_text(encoding="utf-8").split(":")[-1].strip()
    if DATA["excel"]["gamedata_const"]["dataVersion"] != data_version:
        update_data()
    if RESULTS.update:
        update_data()

    print("done.")


def gen_data_paths():
    global EXCEL_DATA_PATH, STORY_DATA_PATH
    global EXCEL_DATA_PATHS, STORY_DATA_PATHS
    global DATA_VERSION_PATH

    EXCEL_DATA_PATH = DATA_PATH / "excel"
    STORY_DATA_PATH = DATA_PATH / "story"
    EXCEL_DATA_PATHS = {
        "activity_table": EXCEL_DATA_PATH / "activity_table.json",
        "gamedata_const": EXCEL_DATA_PATH / "gamedata_const.json",
        "handbook_info_table": EXCEL_DATA_PATH / "handbook_info_table.json",
        "story_review_table": EXCEL_DATA_PATH / "story_review_table.json",
        "story_variables": STORY_DATA_PATH / "story_variables.json",
    }
    STORY_DATA_PATHS = {
        "info": STORY_DATA_PATH / "[uc]info",
        "activities": STORY_DATA_PATH / "activities",
        "obt": STORY_DATA_PATH / "obt",
    }
    DATA_VERSION_PATH = EXCEL_DATA_PATH / "data_version.txt"


def main():
    gen_data_paths()
    load_data()

    if RESULTS.count:
        count_words()
        pickle_path.write_bytes(pickle.dumps(DATA))

    # print_data()
    gen_excel()


if __name__ == "__main__":
    PARSER = argparse.ArgumentParser(
        description="Parse Arknights game data.",
        epilog=r"e.g.: $python %(prog)s data_path",
    )
    PARSER.add_argument(
        "-v",
        "--version",
        action="version",
        version="%(prog)s version 1.0 author:@lengyanyu258 2023年6月17日",
    )
    PARSER.add_argument(
        "data_path", nargs="?", help="Arknights game data directory path."
    )
    PARSER.add_argument(
        "-u",
        "--update",
        action="store_true",
        help="Updating stories.",
    )
    PARSER.add_argument(
        "-c",
        "--count",
        action="store_true",
        help="Counting words.",
    )
    PARSER.add_argument(
        "-s",
        "--style",
        action="store_true",
        help="Setting style.",
    )
    PARSER.add_argument(
        "-ci",
        "--count_info",
        action="store_true",
        help="Counting info words.",
    )
    PARSER.add_argument(
        "-st",
        "--show_total",
        action="store_true",
        help="Show total words.",
    )
    PARSER.add_argument(
        "-sc",
        "--show_counter",
        action="store_true",
        help="Show counter by name.",
    )
    RESULTS = PARSER.parse_args()
    DIR_PATH = RESULTS.data_path
    if RESULTS.data_path:
        if platform.system().lower() == "windows":
            # strip ambiguous chars.
            DIR_PATH = (
                RESULTS.data_path.encode()
                .translate(None, delete='*?"<>|'.encode())
                .decode()
            )
        if (DIR_PATH := pathlib.Path(DIR_PATH)).is_dir():
            DATA_PATH = DIR_PATH
            main()
        else:
            print(f"{os.fspath(DIR_PATH)} is not directory!")
    elif DATA_PATH.is_dir():
        main()
