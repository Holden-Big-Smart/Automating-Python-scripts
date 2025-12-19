import os
import sys

# 配置文件所在的目录
# CONFIG_PATH = os.path.join(
    # os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")),
    # "0_模板文件及初始化",
# )

# 加入系统路径
# sys.path.append(CONFIG_PATH)

# 把配置文件所在目录加入系统路径
sys.path.append(
    os.path.join(
        os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")),
        "0_模板文件及初始化",
    )
)

# 现在导入 config_paths.py
from config_paths import (
    FACEBOOK_PDF_PATH,
    FACEBOOK_OUTPUT_DIR,
    KNIGHT_TEMPLATE_PATH,
    EXCEL_EXPENSE_TEMPLATE,
)

print(
    f"""123FaceBook宣传费的模板文件的路径为：
      【{KNIGHT_TEMPLATE_PATH}】
      【{EXCEL_EXPENSE_TEMPLATE}】"""
)
