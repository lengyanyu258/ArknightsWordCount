import json
import pickle
from argparse import Namespace
from datetime import date
from pathlib import Path

from tqdm import tqdm

from .base import Info
from .count import Count
from .dump import Dump


class GameData(Count, Dump):
    __unknown: dict[str, list[str]] = {"files": [], "commands": [], "heads": []}

    __updated: bool = False
    __counted: bool = False

    def __init__(
        self,
        data_dir_path: str,
        config: Namespace,
        count_config: Namespace,
        dump_config: Namespace,
        args: Namespace,
    ):
        data_dir = Path(data_dir_path)
        if not data_dir.is_dir():
            raise NotADirectoryError(f"{data_dir.absolute()} is not a directory!")

        self.__pickle_file = Path(config.pickle_file_path)
        self.__json_file = Path(config.json_file_path)

        Count.__init__(
            self=self,
            config=count_config,
            unknown=self.__unknown,
            args=args,
        )
        Dump.__init__(
            self=self,
            config=dump_config,
            args=args,
        )

        self.__debug: bool = args.debug

        self.__unknown_files_file = self.__pickle_file.with_name(
            f"{self.__pickle_file.stem}_unknown_files.txt"
        )

        excel_dir = data_dir / "excel"
        self.__story_dir = data_dir / "story"

        self.__data_version_path = excel_dir / "data_version.txt"
        if not self.__data_version_path.exists():
            raise FileNotFoundError(f"{self.__data_version_path.absolute()} not found!")

        self.__excel_dirs: dict[str, Path] = {
            "activity_table": excel_dir / "activity_table.json",
            "gamedata_const": excel_dir / "gamedata_const.json",
            "handbook_info_table": excel_dir / "handbook_info_table.json",
            "story_review_table": excel_dir / "story_review_table.json",
            "story_variables": self.__story_dir / "story_variables.json",
        }
        self.__story_dirs: dict[str, Path] = {
            "info": self.__story_dir / "[uc]info",
            "activities": self.__story_dir / "activities",
            "obt": self.__story_dir / "obt",
        }

        self.__load_data()

    @property
    def version(self) -> tuple[int, ...]:
        return self.__version

    @property
    def date(self) -> date:
        return self.__date

    @property
    def updated(self) -> bool:
        return self.__updated

    @Info("loading...")
    def __load_data(self):
        def parse_version(ver: str):
            return tuple(map(lambda x: int(x), ver.split(".")))

        if self.__pickle_file.exists():
            self.data = pickle.loads(self.__pickle_file.read_bytes())
        elif not self.__pickle_file.parent.exists():
            self.__pickle_file.parent.mkdir(parents=True)

        if self.__json_file.exists():
            self.data.update(json.loads(self.__json_file.read_text(encoding="gb18030")))

        content = self.__data_version_path.read_text(encoding="utf-8")
        self.__version = parse_version(content.split(":")[-1].strip())
        self.__date = date.fromisoformat(content.split()[-2].strip().replace("/", "-"))

        old_version = parse_version(self.data["excel"]["gamedata_const"]["dataVersion"])
        if self.version > old_version:
            info_data = self.data.get("info", {}).get("data", {})
            if self.version > parse_version(info_data.get("数据版本", "0.0.0")):
                self.update()

    @Info("updating story...")
    def __update_story(self):
        info_files = self.__story_dirs["info"].rglob("*.txt")
        activity_files = self.__story_dirs["activities"].rglob("*.txt")
        obt_files = self.__story_dirs["obt"].rglob("*.txt")
        files = list(activity_files) + list(obt_files)
        self.data["story"] = {}

        for info in tqdm(list(info_files), "files"):
            info_relative_path = info.relative_to(self.__story_dirs["info"])
            story_key = info_relative_path.with_suffix("").as_posix()
            file = self.__story_dir / info_relative_path
            if file in files:
                files.remove(file)
                txt = file.read_text(encoding="utf-8")
            else:
                txt = ""
                if self.__debug:
                    # warnings.warn(f"{file} not found!")
                    self.__unknown["files"].append(file.as_posix())
            self.data["story"][story_key] = {
                "info": info.read_text(encoding="utf-8"),
                "txt": txt,
            }

        for file in files:
            file_relative_path = file.relative_to(self.__story_dir)
            story_key = file_relative_path.with_suffix("").as_posix()
            self.data["story"][story_key] = {
                "info": "",
                "txt": file.read_text(encoding="utf-8"),
            }

        if len(self.__unknown["files"]):
            tmp_text = ""
            for i in self.__unknown["files"]:
                tmp_text += f'"{i}",\n'
            self.__unknown_files_file.write_text(tmp_text)

    @Info("updating...")
    def update(self):
        if self.__updated:
            return
        self.__updated = True

        for i in self.__excel_dirs:
            self.data["excel"][i] = json.loads(self.__excel_dirs[i].read_bytes())

        self.__update_story()
        self.count()

    @Info("counting words...")
    def count(self):
        if self.__counted:
            return
        self.__counted = True

        self.count_words()
        self.__pickle_file.write_bytes(pickle.dumps(self.data))

    @Info("start dumping...")
    def dump(self, info: dict) -> Path:
        self.data["info"] = info
        return self.dump_excel()

    @Info("publish file...")
    def publish(self, xlsx_file_path: str, dumped_file: Path):
        import re

        published_file = Path(xlsx_file_path)
        published_dir = published_file.parent
        website_dir = published_dir / "website"

        # remove old xlsx files
        for file_path in published_dir.rglob(f"*{published_file.suffix}"):
            file_path.unlink()

        # add new files
        published_file.unlink(missing_ok=True)
        published_file.hardlink_to(target=dumped_file)
        alternative_file = website_dir / dumped_file.name
        alternative_file.unlink(missing_ok=True)
        alternative_file.hardlink_to(target=dumped_file)

        # modify the index.html file
        index_html_file = website_dir / "index.html"
        index_html_file.write_text(
            re.sub(
                rf"{published_file.stem}_?\d*\{published_file.suffix}",
                alternative_file.name,
                index_html_file.read_text(encoding="utf-8"),
            ),
            encoding="utf-8",
        )

        self.__json_file.write_text(
            json.dumps(
                {"info": self.data["info"]},
                ensure_ascii=False,
                indent=4,
            ),
            encoding="gb18030",
        )
