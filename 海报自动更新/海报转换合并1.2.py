import os
import comtypes.client
from tqdm import tqdm

def merge_ppt_with_template(template_path, input_folder, output_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1  # 必须可见，否则 COM 不允许复制幻灯片

    pptx_files = [
        f for f in os.listdir(input_folder)
        if f.lower().endswith(".pptx") and os.path.isfile(os.path.join(input_folder, f))
    ]

    if not pptx_files:
        print("⚠️ 未找到任何 .pptx 文件。")
        powerpoint.Quit()
        return

    print(f"📂 找到 {len(pptx_files)} 个 PPTX 文件，开始保真合并（使用模板）...\n")

    # ✅ 以模板为基底打开
    dest_ppt = powerpoint.Presentations.Open(template_path, ReadOnly=False, WithWindow=False)

    for file in tqdm(pptx_files, desc="📄 正在追加 Slide"):
        src_path = os.path.join(input_folder, file)
        try:
            src_ppt = powerpoint.Presentations.Open(src_path, ReadOnly=True, WithWindow=False)
            src_ppt.Slides(1).Copy()
            dest_ppt.Slides.Paste()
            src_ppt.Close()
        except Exception as e:
            print(f"❌ 无法复制 {file}：{e}")

    # ✅ 保存为合并文件
    dest_ppt.SaveAs(output_path)
    dest_ppt.Close()
    powerpoint.Quit()
    print(f"\n✅ 已成功合并为：{output_path}")


def main():
    base_dir = os.path.join(os.getcwd(), "海报")
    output_dir = os.path.join(os.getcwd(), "output")
    os.makedirs(output_dir, exist_ok=True)

    # 模板文件路径（放在脚本同级目录）
    template_path = os.path.join(os.getcwd(), "A4纵向模板.pptx")
    if not os.path.exists(template_path):
        print("❌ 未找到模板文件 A4纵向模板.pptx，请确保它与脚本在同一目录下。")
        input("\n按 Enter 键退出...")
        return

    output_path = os.path.join(output_dir, "合并海报.pptx")
    merge_ppt_with_template(template_path, base_dir, output_path)
    input("\n按 Enter 键退出...")


if __name__ == "__main__":
    main()
