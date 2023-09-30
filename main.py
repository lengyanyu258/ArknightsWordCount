import config
from game_data import GameData


def main():
    import datetime
    from pathlib import Path

    game_data = GameData(
        data_dir,
        pickle_file=config.PICKLE_PATH,
        count_config=config.count_config,
        dump_config=config.dump_config,
        args=args,
    )

    if args.update:
        game_data.update()
    if args.count:
        game_data.count()

    data_date = (
        game_data._data_version_path.read_text(encoding="utf-8").split()[-2].strip()
    )
    info = {
        "title": config.info["description"],
        "data": {
            "程序版本": config.info["version"],
            "数据版本": game_data.data["excel"]["gamedata_const"]["dataVersion"],
            "数据日期": data_date.replace("/", "-"),
            "文档日期": f"{datetime.date.today():%Y-%m-%d}",
            # "程序作者": config.info["authors"],
            "程序地址": "https://github.com/lengyanyu258/ArknightsWordCount",
        },
    }
    dump_file = game_data.dump(info)
    if args.publish:
        published_file = Path(config.XLSX_PATH)
        published_file.unlink(True)
        published_file.hardlink_to(dump_file)


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
        "-p",
        "--publish",
        action="store_true",
        help="Save excel file to docs directory.",
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
