import os
import datetime
import glob
import shutil
from pathlib import Path
from PyPDF2 import PdfMerger
import win32com.client
from tqdm import tqdm


def pptx_to_pdf(input_path, output_path):
    """使用 PowerPoint 将 PPTX 转为 PDF"""
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1

    try:
        ppt = powerpoint.Presentations.Open(input_path, WithWindow=False)
        ppt.SaveAs(output_path, 32)  # 32 = PDF
        ppt.Close()
    except Exception as e:
        print(f"\n❌ 转换失败：{input_path}")
        print(e)
    finally:
        powerpoint.Quit()


def main():
    base_dir = Path(__file__).parent
    posters_dir = base_dir / "海报"
    output_dir = base_dir / "output"
    temp_dir = output_dir / "temp_pdf"

    output_dir.mkdir(exist_ok=True)
    temp_dir.mkdir(exist_ok=True)

    # 📌 获取非子目录的 PPTX 文件
    pptx_files = [f for f in posters_dir.glob("*.pptx") if f.is_file()]

    if not pptx_files:
        print("⚠️ 未找到任何 PPTX 文件")
        return

    # 📌 按创建时间从旧到新排序
    pptx_files.sort(key=lambda f: f.stat().st_ctime)

    print(f"📌 共找到 {len(pptx_files)} 个 PPTX 文件，开始转换...\n")

    pdf_list = []

    # === 【阶段 1】PPTX → PDF 转换（含进度条） ===
    for pptx in tqdm(pptx_files, desc="🎨 正在转换为 PDF", ncols=80):
        pdf_path = temp_dir / (pptx.stem + ".pdf")
        pptx_to_pdf(str(pptx), str(pdf_path))

        if pdf_path.exists():
            pdf_list.append(str(pdf_path))
        else:
            print(f"\n❌ 文件转换失败：{pptx.name}")

    if not pdf_list:
        print("❌ 没有成功转换的 PDF 文件，无法合并。")
        return

    # === 【阶段 2】PDF 合并（含进度条） ===
    print("\n📚 正在合并 PDF...\n")
    merger = PdfMerger()

    for pdf in tqdm(pdf_list, desc="📄 合并进度", ncols=80):
        merger.append(pdf)

    # === 【生成最终文件名】 ===
    now = datetime.datetime.now()
    final_name = now.strftime("%Y%m%d_%H%M") + "合并结果.pdf"
    final_output_path = output_dir / final_name

    merger.write(str(final_output_path))
    merger.close()

    # 清理临时目录
    shutil.rmtree(temp_dir)

    print("\n✅ 全部完成！")
    print(f"✨ 输出文件：{final_output_path}\n")


if __name__ == "__main__":
    main()
