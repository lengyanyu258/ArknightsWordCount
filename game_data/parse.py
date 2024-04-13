import collections
import re
import warnings
from argparse import Namespace

from .data import Data


class Parse(Data):
    __ASIDE_NAME = "『旁白』"
    __command_pattern = re.compile(r"\w+")
    __text_pattern = re.compile(r"(?:\[(.+)\])?(.+)?", re.MULTILINE)
    __subtitle_pattern = re.compile(
        r"(<[A-Za-z\d/=#@\.]+>)|({@[Nn]ickname})|(\\r)|(\\n)"
    )
    __line_pattern = re.compile(
        r"([AaPp]\.?[Mm](?!\w)\.?)|(-?[\u03B1-\u03C9\d]+(?:[\.\-:：]\d+)?(?:%|℃|u/L|M)?)|((?:[A-Za-z]+\.(?!\.))+[A-Za-z]*)|([A-Za-z\u03B1-\u03C9\u0400-\u04FF\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u00FF\d]+(?:[/—\-']?[A-Za-z\d]+)*)"
    )

    def __init__(
        self,
        known_commands: list[str],
        unknown: dict[str, list[str]],
        args: Namespace,
    ):
        self.__known_commands: list[str] = known_commands
        self.__unknown_commands: list[str] = unknown["commands"]
        self.__unknown_heads: list[str] = unknown["heads"]

        self.__debug: bool = args.debug
        self.__count_info: bool = args.count_info

    def __parse_line(
        self, command: str, text: str
    ) -> tuple[bool, str, collections.Counter]:
        def get_attribute(cmd_str: str):
            # TODO: use regex
            return cmd_str.split(",")[0].split("=")[1].strip(" '\")")

        # TODO: 使用立绘判断身份
        is_command = True
        control_command = self.__command_pattern.match(command)
        if control_command is None:
            control_command = command
        else:
            control_command = control_command.group()

        if control_command in self.__known_commands or command.startswith("[character"):
            return True, "", collections.Counter()
        elif control_command in (
            "HEADER",
            "Title",
            "Div",
        ):
            return False, "", collections.Counter()
        elif control_command in (
            "Dialog",
            "PopupDialog",
            "VoiceWithin",
            "dialog",
            "warp",
        ):
            # TODO: dialog(head="npc_694_1" 文 activity_table charCardMap
            try:
                head = get_attribute(command)
            except IndexError:
                return True, "", collections.Counter()
            if control_command == "PopupDialog":
                head = self.data["excel"]["story_variables"][head.lstrip("$")]
            if head.startswith("char"):
                try:
                    story_text: str = self.data["excel"]["handbook_info_table"][
                        "handbookDict"
                    ][head]["storyTextAudio"][0]["stories"][0]["storyText"]
                    name = story_text.split("\n")[0].replace("【代号】", "")
                except KeyError:
                    if head in ("char_340_shwazr6"):
                        name = "黑"
                    else:
                        warnings.warn(f"not found {head}")
                        name = head
            else:
                if "head" not in command:
                    name = self.__ASIDE_NAME
                    if (
                        self.__debug
                        and (debug_info := f"no head: {head}")
                        not in self.__unknown_heads
                    ):
                        self.__unknown_heads.append(debug_info)
                else:
                    name = head
                    if self.__debug and head not in self.__unknown_heads:
                        self.__unknown_heads.append(head)
        elif control_command in ("name") or command.startswith("(name"):
            is_command = False
            name = get_attribute(command)
            if name == "":
                name = self.__ASIDE_NAME
        elif control_command in ("Decision", "decision"):
            # is_command = False
            name = "Dr."
            for i in command.split(","):
                if "option" in i:
                    text += "".join(i.split("=")[1].strip(' "').split(";"))
        elif control_command in ("Sticker", "Subtitle"):
            name = self.__ASIDE_NAME
            for i in command.split(","):
                if "text" in i:
                    text += i.split("text=")[1].strip(' "')
        elif control_command in ("narration", "Narration", "isAvatarRight"):
            name = self.__ASIDE_NAME
        elif control_command in ("multiline"):
            name = get_attribute(command)
        else:
            name = ""
            if self.__debug and control_command not in self.__unknown_commands:
                # warnings.warn(f"unknwn command: {command}")
                self.__unknown_commands.append(control_command)

        match_set = set()
        for i in self.__subtitle_pattern.finditer(text):
            match_set.add(i.group())
        for i in match_set:
            # print(i, "|", text)
            text = text.replace(i, " ")
        # if len(match_set):
        #     print(text, "\n")

        words = []
        clean_text = ""
        endpos = 0
        for i in self.__line_pattern.finditer(text):
            word = i.group()
            clean_text += text[endpos : i.start()]
            endpos = i.end()
            words.append(word)
        clean_text += text[endpos:]
        # 方舟特色倒了！狠狠打击水字数 ( ͡• ͜ʖ ͡• ) 标点符号数缩水 38.19%（逃
        # 破案了！(＃°Д°) 原来省略号占了总标点符号数的 45.85%！（现已被削弱为⅛）
        clean_text = clean_text.replace("……", "…").replace("......", "…")
        clean_text = clean_text.replace("——", "—")

        if self.__debug and len(words):
            print(words, command)
            print(text)
            print(clean_text)
            print()
            temp = (
                clean_text.replace(" ", "")
                .replace("/", "")
                .replace(".", "")
                .replace("~", "")
                .replace("—", "")
                .replace("·", "")
                .replace("\\", "")
            )
            if len(temp.encode("utf-8")) % 3 != 0:
                print(
                    f"\n>>> Debug >>>:\n{words} {command}\n{text}\n{temp}\n<<< End <<<\n"
                )

        collection = collections.Counter(words)
        collection.update(clean_text.replace(" ", ""))
        return is_command, name, collection

    def parse_story(self, story: dict):
        command_count = 0
        collection_dict: dict[str, collections.Counter] = {}

        if self.__count_info:
            txts = story.values()
        else:
            txts = (story["txt"],)

        for txt in txts:
            for line in self.__text_pattern.finditer(txt):
                if line.group() == "":
                    continue
                command, text = line.groups()
                if command is None:
                    command = f'name="{self.__ASIDE_NAME}"'
                else:
                    command = command.strip()
                if text is None:
                    text = ""
                else:
                    text = text.strip()
                is_command, name, collection = self.__parse_line(command, text)
                if is_command:
                    command_count += 1
                if collection.total() > 0:
                    if name in collection_dict:
                        collection_dict[name].update(collection)
                    else:
                        collection_dict[name] = collection

        return command_count, collection_dict
