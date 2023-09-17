import config
from game_data import GameData


def main():
    import datetime

    game_data = GameData(
        data_dir,
        pickle_file=config.PICKLE_PATH,
        merge_names=config.dump.merge_names,
        known_commands=config.dump.known_commands,
        args=args,
    )

    if args.update:
        game_data.update()
    if args.count:
        game_data.count()

    info = {
        "title": config.info["description"],
        "data": {
            "程序版本": config.info["version"],
            "数据版本": game_data.data["excel"]["gamedata_const"]["dataVersion"],
            "文档日期": f"{datetime.date.today():%Y-%m-%d}",
            # "程序作者": config.info["authors"],
            "程序地址": "https://github.com/lengyanyu258/ArknightsWordCount",
            "专栏地址": "https://www.bilibili.com/read/cv24618152/",
        },
    }
    game_data.dump(info, config.dump.FONT_NAME)


if __name__ == "__main__":
    import argparse
    import platform

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version="{name} {version} {license} licensed by {authors} 2023年6月17日".format(
            **config.info
        ),
    )
    parser.add_argument(
        "data_dir", nargs="?", help="Arknights game data directory path."
    )

    switch = parser.add_argument_group(title="switch options")
    switch.add_argument(
        "-u",
        "--update",
        action="store_true",
        help="Updating stories.",
    )
    switch.add_argument(
        "-c",
        "--count",
        action="store_true",
        help="Counting words.",
    )
    switch.add_argument(
        "-d",
        "--debug",
        action="store_true",
        help="Show debug info.",
    )
    switch.add_argument(
        "-s",
        "--style",
        action="store_true",
        help="Setting style in excel file.",
    )
    switch.add_argument(
        "-ci",
        "--count_info",
        action="store_true",
        help="Counting info words.",
    )
    switch.add_argument(
        "-st",
        "--show_total",
        action="store_true",
        help="Show total words in csv file.",
    )
    switch.add_argument(
        "-sc",
        "--show_counter",
        action="store_true",
        help="Show counter by name in csv file.",
    )

    parser.usage = "$python %(prog)s [-h] [-v] [{options_title}] [data_dir]".format(
        options_title=switch.title
    )
    parser.description = config.info["description"]
    parser.epilog = "e.g.: $python %(prog)s {data_dir}".format(data_dir=config.DATA_DIR)
    args = parser.parse_args()

    data_dir: str = args.data_dir
    if not data_dir:
        data_dir = config.DATA_DIR

    # strip ambiguous chars.
    data_dir = data_dir.encode().translate(None, delete='*?"<>|'.encode()).decode()

    if platform.system().lower() == "windows":
        import io
        import sys

        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="gb18030")

    main()
