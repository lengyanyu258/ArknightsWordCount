from config import Config, filename


def main():
    import datetime
    import re
    from pathlib import Path

    from game_data import GameData

    game_data = GameData(
        data_dir_path=data_dir_path,
        config=Config.game_data_config,
        count_config=Config.count_config,
        dump_config=Config.dump_config,
        args=args,
    )

    if args.update:
        game_data.update()
    if args.count:
        game_data.count()
    if args.no_dump:
        return

    data_date = game_data._data_version_path.read_text(encoding="utf-8").split()[-2]
    info = {
        "title": Config.info["description"],
        "data": {
            "程序版本": Config.info["version"],
            "数据版本": game_data.data["excel"]["gamedata_const"]["dataVersion"],
            "数据日期": data_date.strip(),
            "文档日期": f"{datetime.date.today():%Y/%m/%d}",
            "文档说明": "https://github.com/lengyanyu258/ArknightsWordCount/wiki",
        },
        "authors": Config.info["authors"],
    }
    dump_file = game_data.dump(info)

    if args.publish:
        # remove old xlsx files
        for file_path in Path(Config.xlsx_file_path).parent.glob("*.xlsx"):
            file_path.unlink()

        # add new files
        published_file = Path(Config.xlsx_file_path)
        published_file.unlink(missing_ok=True)
        published_file.hardlink_to(target=dump_file)
        alternative_file = published_file.with_name(dump_file.name)
        alternative_file.unlink(missing_ok=True)
        alternative_file.hardlink_to(target=dump_file)

        # modify the index.html file
        index_html_file = published_file.with_name("index.html")
        index_html_file.write_text(
            re.sub(
                rf"{filename}_?\d*\.xlsx",
                alternative_file.name,
                index_html_file.read_text(encoding="utf-8"),
            ),
            encoding="utf-8",
        )


if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version="{name} {version}, {license} licensed by {authors} 2023年6月17日".format(
            **Config.info
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
        "-nd",
        "--no_dump",
        action="store_true",
        help="Do not dump data.",
    )

    parser.usage = "python %(prog)s [-h] [-v] [{options_title}] [data_dir]".format(
        options_title=switch.title
    )
    parser.description = Config.info["description"]
    parser.epilog = "e.g.: python %(prog)s -p {data_dir}".format(
        data_dir=Config.DATA_DIR
    )
    args = parser.parse_args()

    data_dir: str = args.data_dir or Config.DATA_DIR

    # strip ambiguous chars.
    data_dir_path = data_dir.encode().translate(None, delete='*?"<>|'.encode()).decode()

    if sys.platform == "win32":
        from io import TextIOWrapper

        sys.stdout = TextIOWrapper(buffer=sys.stdout.buffer, encoding="gb18030")

    main()
