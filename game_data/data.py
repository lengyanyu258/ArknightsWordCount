class Data(object):
    data: dict[str, dict[str, dict]] = {
        "excel": {
            "activity_table": {},
            "gamedata_const": {
                "dataVersion": "0.0.0",
            },
            "story_review_table": {},
        },
        "story": {},
        "count": {"info": {}, "items": {}},
    }

    __indent: int = -1

    def _info(self, info: str, end: bool = False):
        if not end:
            self.__indent += 1

        print("    " * self.__indent + info, flush=True)

        if end:
            self.__indent -= 1
