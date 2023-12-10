# Excel-update

> 一个简单的Python脚本，全程用ChatGPT生成+改报错，最终完美运行
>
> 请注意，<font color='red'>本项目只适合我个人使用，因为代码非常具有定制化，仅供查阅借鉴</font>

# 使用方法

把待转换的xlsx文件命名为1.xlsx，放在跟脚本同级目录下

然后打开控制台输入  `python Excel-update.py` 回车

即可看到成功回显

# 设计背景

作为班里的某班委，需要每周交一次统计表，但是每次改表都需要手动梭哈改日期

我在想能不能用脚本实现自动改日期

但是我又没学多少Python，所以投奔ChatGPT

# 与ChatGPT的对话节选

**问题：** 如何批量让Excel中的日期加一天？

**解答：** 使用 Python 的 `datetime` 模块可以实现日期的加法操作。假设日期在文本中，可以通过正则表达式提取日期并进行操作。代码如下：

```
pythonCopy codeimport re
from datetime import datetime, timedelta

# 读取文本并提取日期
text = "包含日期的文本"

date_pattern = r"(\d{4})年(\d{1,2})月(\d{1,2})日"
matches = re.findall(date_pattern, text)

if matches:
    year, month, day = map(int, matches[0])
    date_obj = datetime(year, month, day)
    new_date = date_obj + timedelta(days=1)
    new_date_str = new_date.strftime("%Y年%m月%d日")
    print("加一天后的日期:", new_date_str)
else:
    print("未找到日期")
```

**问题：** 如果日期在某段文字中怎么办？

**解答：** 你可以尝试提取日期，并对其进行操作，代码示例与之前类似，只需将文本中的日期提取出来即可。

**问题：** 请用 Python 实现这个功能。

**解答：** 可以使用 Pandas 处理 Excel 文件，并结合日期处理模块进行操作。代码如下：

```
pythonCopy codeimport pandas as pd
from datetime import datetime, timedelta

# 读取 Excel 文件
df = pd.read_excel('your_excel_file.xlsx')

# 定义处理日期的函数 increment_date

# 应用函数到日期列 df['DateColumn']

# 将结果保存回 Excel 文件 df.to_excel('output.xlsx', index=False)
```

**问题：** 如果日期混在一段文字里，而不是单独在一个格子里呢？

**解答：** 这需要先提取文本中的日期，然后使用 Python 处理日期。可以使用正则表达式从文本中提取日期，然后进行操作。

（以下对话略去若干次迭代后，代码不断优化和修改）

**问题：** 为什么实际结果中出现了大量的 `None`？

**解答：** 这可能是由于日期格式不匹配导致的。在处理日期时，对日期格式不匹配的文本进行了处理，但可能出现 `None`。

**问题：** 结果仍然出现了大量的 `None`，请优化你的代码。

**解答：** 我们可以增加代码中的容错机制来避免返回 `None`。修改后的代码应该能够更好地处理日期格式不匹配的情况。

# 我的心得体会

只能说，ai太强了，作为新时代的产物，确实给人们带来了极大地方便

告诉他你想干什么，他就能给你生成，太强了