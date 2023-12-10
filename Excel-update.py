import openpyxl
import re
from datetime import datetime, timedelta

# 读取 Excel 文件
file_path = '1.xlsx'  # 修改为你的文件路径
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

# 保存更新后的 Excel 文件
book.save('updated_file.xlsx')

# 成功运行的回显信息
print("*******************************************")
print("**  程序成功运行并完成日期加七天的操作!  **")
print("*******************************************")
print("**  这个世界有了ai将会更加便利——W1ndys   **")
print("**  本项目由ChatGPT强力驱动              **")
print("**  本项目由W1ndys主持构建               **")
print("*******************************************")

