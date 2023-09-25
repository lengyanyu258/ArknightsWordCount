import copy
import collections
from tqdm import tqdm
from pathlib import Path
from argparse import Namespace

from .parse import Parse


class Count(Parse):
    def __init__(
        self,
        config_object,
        unknown: dict[str, list[str]],
        output_template_file: Path,
        args: Namespace,
    ):
        Parse.__init__(
            self,
            known_commands=config_object.known_commands,
            unknown=unknown,
            args=args,
        )

        self.__unknown_commands: list[str] = unknown["commands"]
        self.__unknown_commands_file = output_template_file.with_name(
            f"{output_template_file.stem}_unknown_commands.txt"
        )
        self.__unknown_heads: list[str] = unknown["heads"]
        self.__unknown_heads_file = output_template_file.with_name(
            f"{output_template_file.stem}_unknown_heads.txt"
        )

        from string import punctuation as punc_en
        from zhon.hanzi import punctuation as punc_zh

        self.__punctuation: set[str] = set(punc_en + punc_zh)

    def __count_story(
        self,
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
            for i in self.__punctuation & collection_set:
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

    def count_words(self):
        self.data["count"] = {"info": {}, "items": {}}
        stories = list(self.data["story"].keys())
        for story_id, story in tqdm(
            self.data["excel"]["story_review_table"].items(), "story_review"
        ):
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
                command_count, collection_dict = self.parse_story(
                    self.data["story"][story_key]
                )
                if len(collection_dict) == 0:
                    continue

                entry_type_dict: dict[str, dict] = self.data["count"][
                    "items"
                ].setdefault(entry_type, {"info": {"name": act_type}, "items": {}})
                story_id_dict: dict[str, dict] = entry_type_dict["items"].setdefault(
                    story_id, {"info": {"name": name}, "items": {}}
                )
                story_dict: dict[str, dict] = story_id_dict["items"].setdefault(
                    story_code, {"info": {"name": story_name}, "items": {}}
                )
                avg_dict: dict[str, dict] = story_dict["items"].setdefault(
                    avg_tag, {"info": {}, "items": {}}
                )
                self.__count_story(
                    command_count,
                    collection_dict,
                    [
                        self.data["count"]["info"],
                        entry_type_dict["info"],
                        story_id_dict["info"],
                        story_dict["info"],
                        avg_dict["info"],
                    ],
                )

        basicInfo = self.data["excel"]["activity_table"]["basicInfo"]
        banned_dirname = {"guide", "tutorial", "training", "act1bossrush", "bossrush"}
        for story_key in stories:
            parts = story_key.split("/")
            if banned_dirname & set(parts):
                continue
            command_count, collection_dict = self.parse_story(
                self.data["story"][story_key]
            )
            if len(collection_dict) == 0:
                continue

            # dic = DATA["count"]["items"].setdefault("OTHERS", {"info": {}, "items": {}})
            dic = self.data["count"]
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
            self.__count_story(command_count, collection_dict, dict_list)

        if len(self.__unknown_commands):
            tmp_text = ""
            for i in self.__unknown_commands:
                tmp_text += f'"{i}",\n'
            self.__unknown_commands_file.write_text(tmp_text, encoding="utf-8")

        if len(self.__unknown_heads):
            tmp_text = ""
            for i in self.__unknown_heads:
                tmp_text += f'"{i}",\n'
            self.__unknown_heads_file.write_text(tmp_text, encoding="utf-8")
