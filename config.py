from pathlib import Path as __Path
import tomllib as __tomllib

__pyproject: dict[str, dict] = __tomllib.loads(
    __Path(__file__).with_name("pyproject.toml").read_text(encoding="utf-8")
)

info: dict = __pyproject["tool"]["poetry"]
DATA_DIR = r"./Github/ArknightsGameData/zh_CN/gamedata"
PICKLE_PATH = r"./tmp/Arknights_Word_Count.pkl"
XLSX_PATH = r"./docs/Arknights_Word_Count.xlsx"


class Dump:
    FONT_NAME = "Sarasa Mono Slab SC"

    name_prefix = ["发言人", "审判官", "大审判官", "无言的", "小"]
    name_suffix = ["骑士", "？"]
    erase_names = ["小黑", "教宗骑士", "感染者骑士"]
    merge_names = [
        ("？？？？", "？？？", "？？", "？"),
        ("“焰尾”索娜", "“焰尾”骑士", "索娜"),
        ("伊万杰利斯塔十一世", "教宗"),
        ("博士", "Dr."),
        ("埃尼斯", "苍苔"),  # 代号尚未在剧情中出现
        ("塞诺蜜", "砾"),
        ("审判官艾丽妮", "艾丽妮"),
        ("微光守夜人", "不叫微光守夜人的黎博利", "菲亚梅塔"),
        ("恩希欧迪斯", "银灰"),
        ("玛嘉烈", "临光"),
        ("玛莉娅", "瑕光"),
        ("祖玛玛", "森蚺"),
        ("阿黛尔", "艾雅法拉"),
        ("陈晖洁", "陈"),
    ]


class Count:
    known_commands = [
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
    ]
