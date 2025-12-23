# config_paths.py
import os

# ========== 项目根目录 ==========
# 获取当前项目的根目录的绝对路径(测试整理目录)
ROOT_DIR = os.path.abspath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
)
# print(f"当前项目的根目录为：\n【{ROOT_DIR}】")
# 【d:\Desktop\SK\sync\脚本\行政清单领款单相关\测试整理目录】

# ========== 课程行政清单领款单脚本所需路径 ==========
# 当前脚本所在目录
# BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# print(BASE_DIR)
# 测试整理目录\0_模板文件及初始化
# => 测试整理目录\1A_课程行政清单_领款单

# 输入数据文件
course_file = os.path.join(
    ROOT_DIR, "1A_课程行政清单_领款单", "屯門婦聯 - 會員及課程管理系統 - 課程.csv"
)
# print("course_file:", course_file)
# 测试整理目录\1A_课程行政清单_领款单\屯門婦聯 - 會員及課程管理系統 - 課程.csv
receipt_file = os.path.join(
    ROOT_DIR, "1A_课程行政清单_领款单", "屯門婦聯 - 會員及課程管理系統 - 課程收據.csv"
)

template_path = os.path.join(ROOT_DIR, "0_模板文件及初始化", "行政清单-模板.docx")
# print("template_path:", template_path)
# 测试整理目录\0_模板文件及初始化\行政清单-模板.docx
template_lingkuan_path = os.path.join(
    ROOT_DIR, "0_模板文件及初始化", "领款单-模板.docx"
)
# print("template_lingkuan_path:", template_lingkuan_path)
# 测试整理目录\0_模板文件及初始化\领款单-模板.docx
template_output_spending_path = os.path.join(
    ROOT_DIR,
    "2_Excel滙入記錄模板-支出賬文件",
    "屯門婦聯 - 會計及財務記賬系統-Excel滙入記錄模板-支出賬.xlsx",
)
# print("template_output_spending_path:", template_output_spending_path)
# 测试整理目录\2_Excel滙入記錄模板-支出賬文件\屯門婦聯 - 會計及財務 記賬系統-Excel滙入記錄模板-支出賬.xlsx

# 输出目录
# output_dir = os.path.join(BASE_DIR, "output")
output_dir = os.path.join(ROOT_DIR, "1A_课程行政清单_领款单", "output")
# print("output_dir:", output_dir)
# 测试整理目录\1A_课程行政清单_领款单\output

# 如果 output_dir 文件夹不存在，就自动创建它；如果已存在，也不会报错。
# os.makedirs(output_dir, exist_ok=True)

# 历史清单文件（将自动生成或追加）
excel_path = os.path.join(output_dir, "历史清单汇总.xlsx")

# Excel汇入记录模板文件路径（支出賬）
template_output_spending_actual_path = os.path.join(
    ROOT_DIR,
    "2_Excel滙入記錄模板-支出賬文件",
    "屯門婦聯 - 會計及財務記賬系統-Excel滙入記錄模板-支出賬.xlsx",
)


# ========== FaceBook脚本所需路径 ==========
# FaceBook宣传费的读取文件的路径(测试整理目录\FaceBook宣传费\input\invoices.pdf)
FACEBOOK_PDF_PATH = os.path.join(
    ROOT_DIR, "1C_FaceBook宣传费领款单", "FaceBook宣传费文件", "invoices.pdf"
)
# print(f"FaceBook读取路径为：\n【{FACEBOOK_PDF_PATH}】")

# FaceBook宣传费的输出文件的路径(测试整理目录\FaceBook宣传费\output)
FACEBOOK_OUTPUT_DIR = os.path.join(ROOT_DIR, "1C_FaceBook宣传费领款单", "output")
# print(f"FaceBook输出路径为：\n【{FACEBOOK_OUTPUT_DIR}】")

# FaceBook宣传费的模板文件的路径(测试整理目录\FaceBook宣传费\output)
KNIGHT_TEMPLATE_PATH = os.path.join(
    ROOT_DIR, "0_模板文件及初始化", "Knight Creative Limited模板.docx"
)
# print(f"FaceBook模板路径为：\n【{KNIGHT_TEMPLATE_PATH}】")


# ========== Excel 会计系统支出账的读取文件 ==========
EXCEL_EXPENSE_TEMPLATE = os.path.join(
    ROOT_DIR,
    "2_Excel滙入記錄模板-支出賬文件",
    "屯門婦聯 - 會計及財務記賬系統-Excel滙入記錄模板-支出賬.xlsx",
)
# print(f"模板路径为：\n【{EXCEL_EXPENSE_TEMPLATE}】")

# # ========== 你可继续在此新增更多路径（行政清单、其他领款单等等） ==========
