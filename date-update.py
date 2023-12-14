import os
import openpyxl
import re
from datetime import datetime, timedelta

# 定义日期提取和处理函数
def increment_date(text):
    date_pattern = r"(\d{4})年(\d{1,2})月(\d{1,2})日"
    matches = re.findall(date_pattern, text)
    if matches:
        year, month, day = map(int, matches[0])
        try:
            date_obj = datetime(year, month, day)
            new_date = date_obj + timedelta(weeks=X)  # X 是代表周数的整数变量
            new_date_str = new_date.strftime("%Y年%m月%d日")
            return re.sub(date_pattern, new_date_str, text)
        except ValueError:
            return text
    else:
        return text

# 代表周数的整数变量
X = 5  # 设置为你希望的周数

# 定义文件名匹配模板
file_name_template = r"22级网络空间安全班第(\d+)周手机入袋情况.xlsx"

# 查找匹配文件名模板的文件
matching_files = []
for filename in os.listdir('.'):  # 切换为文件所在目录
    if re.match(file_name_template, filename):
        matching_files.append(filename)

if len(matching_files) != 1:
    print("找到多个符合条件的文件或未找到文件，请检查文件夹内容。")
else:
    # 获取唯一匹配的文件名
    file_name = matching_files[0]

    # 读取 Excel 文件
    file_path = file_name
    sheet_name = 'Sheet1'  # 修改为你的表格名

    # 读取 Excel 文件
    book = openpyxl.load_workbook(file_path)
    sheet = book[sheet_name]

    # 行号列表（1, 12, 25, 37, 49）包含日期的行
    target_rows = [1, 13, 25, 37, 49]

    # 定义日期提取和处理函数
    def increment_date(text):
        date_pattern = r"(\d{4})年(\d{1,2})月(\d{1,2})日"
        matches = re.findall(date_pattern, text)
        if matches:
            year, month, day = map(int, matches[0])
            try:
                date_obj = datetime(year, month, day)
                new_date = date_obj + timedelta(days=7)
                new_date_str = new_date.strftime("%Y年%m月%d日")
                return re.sub(date_pattern, new_date_str, text)
            except ValueError:
                return text
        else:
            return text

    # 遍历指定行，并对包含日期的文本进行处理
    for row_num in target_rows:
        cell_range = sheet[f"A{row_num}:O{row_num}"]
        for row in cell_range:
            for cell in row:
                if cell.value:
                    cell_value = str(cell.value)
                    new_value = increment_date(cell_value)
                    if new_value != cell_value:
                        cell.value = new_value

    # 保存更新后的 Excel 文件（覆盖原始文件）
    book.save(file_path)

    # 成功运行的回显信息
    print("\n")
    print("*******************************************")
    print("**  程序成功运行并完成日期加七天的操作!  **")
    print("*******************************************")
    print("**  这个世界有了ai将会更加便利——W1ndys   **")
    print("**  本项目由ChatGPT强力驱动              **")
    print("**  本项目由W1ndys主持构建               **")
    print("*******************************************")
