# ArknightsWordCount

《明日方舟》字词统计

## 说明

详见：[Wiki](https://github.com/lengyanyu258/ArknightsWordCount/wiki).

## 使用

需安装 Python 包管理器：[Poetry](https://python-poetry.org/docs/#installation)

```powershell
git clone --depth=1 "https://github.com/lengyanyu258/ArknightsWordCount.git"

cd ArknightsWordCount

poetry shell

poetry install

python main.py -h
```

## 其他

[配置文件（config.py）](https://github.com/lengyanyu258/ArknightsWordCount/blob/main/config.py)中的 `merge_names` 记录的是同一人在文本中出现过的不同名称，比如：

> 微光守夜人 -> 不叫微光守夜人的黎博利 -> 菲亚梅塔
>
> 恩希欧迪斯 -> 银灰
>
> 阿黛尔 -> 艾雅法拉

目前该内容尚未收录全，仍需要建设，欢迎大家[参与](https://github.com/lengyanyu258/ArknightsWordCount/edit/main/config.py)！
