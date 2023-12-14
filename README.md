# Excel-date-update

> 一个简单的Python脚本，全程用ChatGPT生成+改报错，最终完美运行
>
> 请注意，<font color='red'>本项目只适合我个人使用，因为代码非常具有定制化，仅供查阅借鉴</font>

# 更新日志

2023年12月14日13:09:07更新了自动扫描文件目录，直接替换源文件且自动扫描目录中命名为：“22级网络空间安全班第X周手机入袋情况”的文件

2023年12月11日22:46:08发布了release

2023年12月11日19:42:31修改了文件名字和一些细节

2023年12月11日13:45:35增加了一个批处理文件，更加懒了，原来需要输入命令，现在只需要双击bat文件就可以实现了，呜呜呜懒人有懒福

2023年12月10日23:12:56纠正了增加七天的bug（原来设计的是增加1天

# 使用方法

首先运行下面安装必备的库

```
pip install pandas openpyxl xlrd
```

然后[点击这里](https://github.com/W1ndys/Excel-date-update/releases/download/Excel-date-update/Excel-date-update_v1.0.zip)下载必要的文件

你的Excel文件名字应该是XX级XX班第X周手机入袋情况

所以打开文件 `date-update.py`

修改第26行代码

```
file_name_template = r"22级网络空间安全班第(\d+)周手机入袋情况.xlsx"
```

把这个名字改为适合你的情况的名字，例如99级一班第(\d+)周情况.xlsx。

注意：这个名字除了(\d+)其余部分均可修改，只需要改为待转换文件名一致就可以，(\d+)是整形数字的通配符

然后双击：一键执行.bat就可以了

即可看到成功回显

# 核心代码解释

代码第14行，因为我的表格是这些行，所以就建立了这么个数组

```
# 行号列表（1, 12, 25, 37, 49）包含日期的行
target_rows = [1, 13, 25, 37, 49]
```

是需要修改的行数，根据自身需要可以自行更改

# 设计背景

作为某班委，需要每周交一次统计表，但是每次改表都需要手动梭哈改五遍日期

我在想能不能用脚本或者函数实现一键改日期

期初想过用Excel的格式化，但是没找到合适的，因为文档格式是合并单元格，没法一键替换

于是我又转向了Python

但是我又没学多少Python，所以投奔ChatGP

首先我问Python能不能实现

然后把ChatGP给我的代码一遍遍在自己电脑上跑

遇到报错就扔给chatGPT让他自己改

终于改到没有报错之后运行了一下

看到生成的文件中有大量的杂乱信息

于是我又问ChatGPT，但他给我的解释是字符

这又不好解决，所以我采用了另一种方法

既然我需要改的日期文本在固定的行数

我直接改那几行就好了，其他的内容不需要扫描

所以干脆让ai写个只改几行的

果不其然

成功了！

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

# 运行环境

Python

openpyxl库

