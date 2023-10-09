# ArknightsWordCount

《明日方舟》字词统计

## 说明

由于采取了分别统计字词数（汉字字数与单词数量）和标点符号数的方式，所以为了与其他人做出的统计字数进行区分，这里采用仅以字词数（不计入标点符号数）作为基准重新排序。

这种将字词数单独统计出来有一个优势就是，我们可以直接根据字词量来预估剧情的观看时长。直播时长大约是每 1 万字 1 小时（包含打关时间，不算发癫时间），比如《孤星》12 万字大约需要直播 12 小时。如果选择不念台词、自己默读，则时长减半。

### 注意事项

1. 包含『情报处理室』中『公共事务实录』与『特别行动记述』所有内容；

2. 包含主题曲内笔记文本（如第 10 章《破碎日冕》）、别传中关卡外文本（如《长夜临光》、《叙拉古人》）；

3. 包含愚人节等临时活动文本、肉鸽文本、干员密录等文本；

4. 为了使问题简化，目前**不**包含人员档案、训练关卡与游戏内引导等文本。

5. 每个单词、代号、数字等（如 `Dijkstra`、`U-Official`、`6152.31`）算作一个字词处理；

6. 将省略号 `......` 六个字符统一换算作 `……` 两个字符处理；

7. **统计结果并不精确，仍有改进空间，切勿全信。**

### 理想案例

- （我是主播）今天打算直播主线第十三章！

  - 让我来看看这一章节的文本量是多少——

    - 噢！竟然有 11.8 万字！看来得直播 11.8 小时才行（打关时间等包含在内）！这简直是太～久～了！

    - 我每天最多只有 4 小时的时间，不得已得把它拆分成 3 天。

  - 让我来看看应该怎么拆分——

    - 切换到底部的 `Simple` 表单（或是 `MAINLINE` 表单），看一下每关大概有多少字（换算成时间），仔细盘算分配一下直播内容！

    - 稍作考虑，我打算第一天直播 `13-1`～`13-5` 的内容，第二天直播 `13-6`～`13-15`，第三天为 `13-16`～`13-22`，完美！

- （我是观众）今天打算观看主线十三章的直播！

  - 让我来看看一共有多少字——

  - OMG，这要是看直播的话，得连看将近 12 小时！下午 4 点开播的话，这得一股脑干到凌晨 4 点！

  - 而我明天还有早八，决定了，今晚跟一半剩下的看录播！（谋定而后动，不再迷茫！）

- 其他，自己用：

  - 方舟出新活动了！我来看看这次新主线剧情有多少万字？——OMG，算上主题曲内的笔记记录文本，一共得有 12 万字！自己看完得花 6 小时！今晚的时间安排心里有数儿了！

  - 今天终于有空了，打算补一补以前落下的剧情！让我来看看这个活动的文本量有多少——啊？《生于黑夜》只有 2.9 万字？这岂不是只需沉浸一个半小时即可看完？之前还以为剧情很长呢一直不敢起个头看，这下心里有谱儿了！

  - ……

- 其他，纯好奇：

  - 噢！～方舟文本量原来是这么多！

  - 嗷！～原来像《长夜临光》、《孤星》等剧情的文本字数竟然是这些！

  - 啊！～原来博士说了这么多的话！

  - ……

- ……

## 使用

需安装 Python 包管理器：[Poetry](https://python-poetry.org/docs/#installation)

```powershell
git clone --depth=1 "https://github.com/lengyanyu258/ArknightsWordCount.git"

cd .\ArknightsWordCount\

poetry shell

poetry install

python .\main.py -h
```

## 其他

[配置文件（config.py）](https://github.com/lengyanyu258/ArknightsWordCount/blob/main/config.py)中的 `merge_names` 记录的是同一人在文本中出现过的不同名称，比如：

> 微光守夜人 -> 不叫微光守夜人的黎博利 -> 菲亚梅塔
>
> 恩希欧迪斯 -> 银灰
>
> 阿黛尔 -> 艾雅法拉

目前该内容尚未收录全，仍需要建设，欢迎大家[参与](https://github.com/lengyanyu258/ArknightsWordCount/edit/main/config.py)！
