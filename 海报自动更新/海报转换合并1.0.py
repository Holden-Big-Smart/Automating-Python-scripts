import os
import datetime
import glob
import shutil
from pathlib import Path
from PyPDF2 import PdfMerger
import win32com.client


def pptx_to_pdf(input_path, output_path):
    """使用 PowerPoint 将 pptx 转换为 pdf"""
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1

    try:
        ppt = powerpoint.Presentations.Open(input_path, WithWindow=False)
        ppt.SaveAs(output_path, 32)  # 32 = PDF 格式
        ppt.Close()
    except Exception as e:
        print(f"❌ 转换失败：{input_path}")
        print(e)
    finally:
        powerpoint.Quit()


def main():
    base_dir = Path(__file__).parent
    posters_dir = base_dir / "海报"
    output_dir = base_dir / "output"
    temp_dir = base_dir / "output" / "temp_pdf"

    output_dir.mkdir(exist_ok=True)
    temp_dir.mkdir(exist_ok=True)

    # 📌 获取海报目录下的所有 .pptx（不含子目录）
    pptx_files = [
        f for f in posters_dir.glob("*.pptx")
        if f.is_file()
    ]

    if not pptx_files:
        print("⚠️ 未在『海报』文件夹中找到 PPTX 文件")
        return

    # 📌 按创建时间排序（从旧到新）
    pptx_files.sort(key=lambda f: f.stat().st_ctime)

    print("📌 共找到 PPTX 文件：", len(pptx_files))

    pdf_list = []

    # 📌 逐个转换为 PDF
    for pptx in pptx_files:
        pdf_path = temp_dir / (pptx.stem + ".pdf")
        print(f"👉 正在转换：{pptx.name} → {pdf_path.name}")
        pptx_to_pdf(str(pptx), str(pdf_path))

        if pdf_path.exists():
            pdf_list.append(str(pdf_path))
        else:
            print(f"❌ 转换失败（文件不存在）：{pptx.name}")

    # 📌 合并 PDF
    if not pdf_list:
        print("❌ 没有 PDF 可以合并")
        return

    merger = PdfMerger()

    for pdf in pdf_list:
        merger.append(pdf)

    # 📌 生成最终文件名
    now = datetime.datetime.now()
    final_name = now.strftime("%Y%m%d_%H%M") + "合并结果.pdf"
    final_output_path = output_dir / final_name

    merger.write(str(final_output_path))
    merger.close()

    # 📌 清理临时 PDF
    shutil.rmtree(temp_dir)

    print("✅ 合并完成！")
    print("✨ 输出文件：", final_output_path)


if __name__ == "__main__":
    main()
