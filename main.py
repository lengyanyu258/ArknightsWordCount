from config import Config
from game_data import GameData


def manipulate(game_data: GameData):
    from datetime import datetime, timedelta, timezone

    # Used by GitHub Actions
    if args.auto_update:
        if not game_data.need_update:
            print("No need to update!")

        if args.test_update:
            import os

            # 设置环境变量以供 GitHub Actions 捕获
            with open(os.environ["GITHUB_OUTPUT"], "a") as github_output:
                print(f"test_update={game_data.need_update}", file=github_output)

    if args.no_dump:
        return

    if args.update or game_data.need_update:
        game_data.update()
    if args.count:
        game_data.count()

    datetime_now = datetime.now(timezone(timedelta(hours=8)))
    info = {
        "title": Config.info["description"],
        "data": {
            "程序版本": Config.info["version"],
            "数据版本": ".".join(map(lambda x: str(x), game_data.version)),
            "数据日期": game_data.date.isoformat(),
            "文档日期": datetime_now.date().isoformat(),
            "文档说明": "https://github.com/lengyanyu258/ArknightsWordCount/wiki",
        },
        "authors": Config.info["authors"],
    }
    if args.auto_update:
        info["data"][
            "其他说明"
        ] = f"{datetime_now.time().isoformat(timespec='seconds')} 自动更新"
    dumped_file = game_data.dump(info)

    if args.publish:
        game_data.publish(Config.xlsx_file_path, dumped_file)


def main():
    data_dir_set = {data_dir_path}
    if args.all:
        data_dir_set.update({*Config.DATA_DIRS})

    game_data_objs: list[GameData] = []
    for data_dir in data_dir_set:
        game_data_objs.append(
            GameData(
                data_dir_path=data_dir,
                config=Config.game_data_config,
                count_config=Config.count_config,
                dump_config=Config.dump_config,
                args=args,
            )
        )
    game_data_objs.reverse()
    for game_data in game_data_objs:
        if game_data.need_update:
            break
    else:
        game_data = game_data_objs[-1]

    manipulate(game_data)


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
        help="Save excel file to docs directory & Publish website files.",
    )
    switch.add_argument(
        "-a",
        "--all",
        action="store_true",
        help="Try to update all DATA_DIRS & Publish it.",
    )
    switch.add_argument(
        "--test_update",
        action="store_true",
        help="Test Update by GitHub Action flag.",
    )
    switch.add_argument(
        "--auto_update",
        action="store_true",
        help="Auto Update by GitHub Action flag.",
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
        data_dir=Config.DATA_DIRS[0]
    )
    args = parser.parse_args()

    data_dir: str = args.data_dir or Config.DATA_DIRS[0]

    # strip ambiguous chars.
    data_dir_path = data_dir.encode().translate(None, delete='*?"<>|'.encode()).decode()

    if sys.stdout.encoding == "gbk":
        from io import TextIOWrapper

        sys.stdout = TextIOWrapper(buffer=sys.stdout.buffer, encoding="gb18030")

    main()
