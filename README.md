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

5. **统计结果并不精确，仍有改进空间，切勿全信。**

## 使用

需安装 Python 包管理器：[Poetry](https://python-poetry.org/docs/#installation)

```powershell
git clone --depth=1 https://github.com/lengyanyu258/ArknightsWordCount.git

cd .\ArknightsWordCount\

poetry shell

poetry install

python .\main.py -h
```

## 贡献

[`config.py`](https://github.com/lengyanyu258/ArknightsWordCount/blob/main/config.py) 中的 `merge_names` 内容仍需要建设，欢迎大家参与！

## 其他

- 获取 Excel 表格：

  蓝奏云（无需登录注册，直接下载）网盘文件夹地址：
  
  链接：<https://lengyanyu258.lanzoub.com/b04winecj>
  
  密码：awc
