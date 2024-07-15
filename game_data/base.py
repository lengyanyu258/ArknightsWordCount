from functools import wraps
from typing import Callable


class Base(object):
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


class Info:
    # 这个字典将在所有被装饰的函数间共享
    _shared_data = {"indent": -1}

    def __init__(self, message: str):
        self.msg = message

    def __call__(self, func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            self.log_start(self.msg)

            result = func(*args, **kwargs)

            self.log_stop(f"{('D', 'd')[self.msg[0].islower()]}one.")

            return result

        return wrapper

    def log_start(self, info: str):
        self._shared_data["indent"] += 1
        self.log(info)

    def log(self, info: str):
        print("    " * self._shared_data["indent"] + info, flush=True)

    def log_stop(self, info: str):
        self.log(info)
        self._shared_data["indent"] -= 1
