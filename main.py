from config import Config
from game_data import GameData


def manipulate(game: GameData):
    from datetime import datetime, timedelta, timezone

    print("Current GameData Dir:", game.data_dir)

    # Used by GitHub Actions
    if args.test_update:
        import os

        # 设置环境变量以供 GitHub Actions 捕获
        # 如果是手动执行，则会强制更新：need_update = github.event_name != 'schedule' || test_update
        with open(os.environ["GITHUB_OUTPUT"], "a") as github_output:
            print(
                f"test_update={str(game.need_update).lower()}",
                file=github_output,
            )

        return

    datetime_now = datetime.now(timezone(timedelta(hours=8)))
    game.data["info"] = {
        "title": Config.info["description"],
        "data": {
            "程序版本": Config.info["version"],
            "数据版本": ".".join(map(lambda x: str(x), game.version)),
            # "数据日期": game.date.isoformat(),
            "文档日期": datetime_now.date().isoformat(),
            "文档说明": "https://github.com/lengyanyu258/ArknightsWordCount/wiki",
        },
        "authors": Config.info["authors"],
    }
    if args.update or game.need_update or args.auto_update:
        game.update()
    if args.count:
        game.count()
    if args.no_dump:
        print("No dump file will be generated.")
        return

    if args.auto_update:
        import os

        # 读取环境变量
        event_name = os.getenv("GITHUB_EVENT_NAME", "unknown")

        # 根据触发事件打印不同的消息
        other_info = datetime_now.time().isoformat(timespec="seconds")
        if event_name == "workflow_dispatch":
            other_info = f"{other_info} 手动更新"
        elif event_name in ["push", "pull_request", "schedule"]:
            other_info = f"{other_info} 自动更新"
        else:
            other_info = f"{other_info} update by {event_name}"

        game.data["info"]["data"]["其他说明"] = other_info

    dumped_file = game.dump()

    if args.publish:
        game.publish(Config.xlsx_file_path, dumped_file)


def main():
    data_dir_set = {data_dir_path}
    if args.all:
        data_dir_set.update({*Config.DATA_DIRS})

    game_data_objs: list[GameData] = []
    for data_dir in data_dir_set:
        try:
            game_data_objs.append(
                GameData(
                    data_dir_path=data_dir,
                    config=Config.game_data_config,
                    count_config=Config.count_config,
                    dump_config=Config.dump_config,
                    args=args,
                )
            )
        except NotADirectoryError as e:
            print(f"{data_dir} 数据目录不存在或无法读取: {e}")
            continue
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
