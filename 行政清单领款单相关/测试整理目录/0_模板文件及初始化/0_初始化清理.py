import os
import shutil

# ========== 获取根目录（脚本上上级目录）==========
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))  # 上上级目录作为项目根目录

# ========== 路径定义 ==========

folders_to_clear = [
    os.path.join(BASE_DIR, "1A_课程行政清单_领款单", "output"),
    os.path.join(BASE_DIR, "2_Excel滙入記錄模板-支出賬文件"),
    os.path.join(BASE_DIR, "3_會計及財務記賬系統 - 下載文件"),
]

files_to_copy = [
    {
        "src": os.path.join(BASE_DIR, "0_模板文件及初始化", "历史清单汇总.xlsx"),
        "dst": os.path.join(
            BASE_DIR, "1A_课程行政清单_领款单", "output", "历史清单汇总.xlsx"
        ),
    },
    {
        "src": os.path.join(
            BASE_DIR,
            "0_模板文件及初始化",
            "屯門婦聯 - 會計及財務記賬系統-Excel滙入記錄模板-支出賬.xlsx",
        ),
        "dst": os.path.join(
            BASE_DIR,
            "2_Excel滙入記錄模板-支出賬文件",
            "屯門婦聯 - 會計及財務記賬系統-Excel滙入記錄模板-支出賬.xlsx",
        ),
    },
]

# ========== 执行清理 ==========

print("🧹 初始化清理开始...\n")

for folder in folders_to_clear:
    print(f"📁 清理目录：{folder}")
    if os.path.exists(folder):
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                    print(f"  ⛔ 已删除文件：{file_path}")
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                    print(f"  ⛔ 已删除文件夹：{file_path}")
            except Exception as e:
                print(f"  ⚠️ 删除失败：{file_path}，原因：{e}")
    else:
        print(f"  ⚠️ 目录不存在，已跳过")

# ========== 执行复制 ==========

print("\n📋 正在复制模板文件...\n")

for file in files_to_copy:
    src = file["src"]
    dst = file["dst"]
    try:
        dst_folder = os.path.dirname(dst)
        os.makedirs(dst_folder, exist_ok=True)
        shutil.copy2(src, dst)
        print(f"✅ 已复制模板文件：{src} ➜ {dst}")
    except Exception as e:
        print(f"❌ 复制失败：{src} ➜ {dst}，原因：{e}")

print("\n✅ 初始化完成！")
input("📌 按回车键退出 ...")
