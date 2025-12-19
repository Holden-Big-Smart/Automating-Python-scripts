import os
from moviepy import VideoFileClip, concatenate_videoclips

def main():
    input_folder = "videos"
    output_folder = "output"
    output_filename = "output.mp4"

    # 确保输出文件夹存在pip install moviepy

    os.makedirs(output_folder, exist_ok=True)

    # 获取所有视频文件并按文件名排序
    video_files = sorted([
        f for f in os.listdir(input_folder)
        if f.lower().endswith((".mp4", ".mov", ".avi", ".mkv"))
    ])

    if not video_files:
        print("未找到任何视频文件")
        return

    clips = []
    for filename in video_files:
        filepath = os.path.join(input_folder, filename)
        try:
            clip = VideoFileClip(filepath)
            clips.append(clip)
            print(f"已加载：{filename}")
        except Exception as e:
            print(f"跳过无法加载的视频文件：{filename}，错误：{e}")

    if not clips:
        print("所有视频都无法加载，退出。")
        return

    # 拼接视频
    final_clip = concatenate_videoclips(clips, method="compose")

    # 输出路径
    output_path = os.path.join(output_folder, output_filename)
    final_clip.write_videofile(output_path, codec="libx264", audio_codec="aac")

    # 清理资源
    for clip in clips:
        clip.close()
    final_clip.close()

    print(f"视频已输出到：{output_path}")

if __name__ == "__main__":
    main()
