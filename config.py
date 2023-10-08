import tomllib
from typing import Any
from pathlib import Path
from argparse import Namespace

pyproject: dict[str, dict[str, Any]] = tomllib.loads(
    Path(__file__).with_name("pyproject.toml").read_text(encoding="utf-8")
)
filename: str = pyproject["tool"]["poetry"]["name"]


class Config:
    DATA_DIR = r"./Github/ArknightsGameData/zh_CN/gamedata"

    info: dict[str, Any] = pyproject["tool"]["poetry"]
    xlsx_file_path = f"./docs/{filename}.xlsx"

    game_data_config = Namespace(pickle_file_path=f"./tmp/{filename}.pkl")

    dump_config = Namespace(
        FONT_NAME="Sarasa Mono Slab SC",
        output_file_path=game_data_config.pickle_file_path,
        # 名称前缀
        name_prefix=["发言人", "审判官", "大审判官", "小"],
        # 名称后缀
        name_suffix=["骑士", "？"],
        # 排除合并后的名称
        erase_names=["小黑", "小游客", "小村民", "教宗骑士", "感染者骑士"],
        # 合并名称
        merge_names=[
            ("？？？？？", "？？？", "？"),
            ("“焰尾”索娜", "“焰尾”", "焰尾", "索娜"),
            ("伊万杰利斯塔十一世", "教宗"),
            ("博士", "Dr."),
            ("埃内斯托", "龙舌兰"),
            ("埃尼斯", "苍苔"),  # 代号尚未在剧情中出现
            ("塞诺蜜", "砾"),
            ("娜塔莉娅", "早露"),
            ("微光守夜人", "不叫微光守夜人的黎博利", "苦难陈述者", "菲亚梅塔"),
            ("恩希欧迪斯", "银灰"),
            ("拉维妮娅", "斥罪"),
            ("无言的达里奥", "达里奥"),
            ("玛嘉烈", "临光"),
            ("玛莉娅", "瑕光"),
            ("祖玛玛", "森蚺"),
            ("艾沃娜", "野鬃"),
            ("莱昂图索", "伺夜"),
            ("费德里科", "送葬人"),
            ("费斯特", "白铁"),
            ("里凯莱", "隐现"),
            ("阿芙朵嘉", "鸿雪"),
            ("阿赫茉妮", "和弦"),
            ("阿黛尔", "艾雅法拉"),
            ("陈晖洁", "陈"),
        ],
    )

    count_config = Namespace(
        output_file_path=game_data_config.pickle_file_path,
        # TODO: use lower case
        known_commands=[
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
            "cgitem",
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
            "hidecgitem",
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
            "timerclear",
            "timersticker",
            "Tutorial",
            "verticalbg",
            "Video",
            "withdraw",
        ],
    )
