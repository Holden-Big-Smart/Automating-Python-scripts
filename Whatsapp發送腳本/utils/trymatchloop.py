# 文件名：循环截图工具.py

import pyautogui
import cv2
import numpy as np
import time

def 循环截图匹配(左上角x, 左上角y, 宽, 高, 匹配路径, 匹配精确度=0.85, 循环间隔=1):
    """
    在指定区域内持续截图匹配，直到匹配成功，返回目标图像中心坐标。

    :param 左上角x: 截图区域左上角 X 坐标
    :param 左上角y: 截图区域左上角 Y 坐标
    :param 宽: 区域宽度
    :param 高: 区域高度
    :param 匹配路径: 模板图像路径
    :param 匹配精确度: 匹配阈值（0~1），默认 0.85
    :param 循环间隔: 每次识图之间的间隔时间（秒）
    :return: (x, y) 目标图像中心坐标
    """
    print(f"[等待识图] 模板：{匹配路径} | 区域：({左上角x}, {左上角y}, {宽}, {高})")

    while True:
        # 局部截图
        局部截图图片 = pyautogui.screenshot(region=(左上角x, 左上角y, 宽, 高))
        局部截图图片 = cv2.cvtColor(np.array(局部截图图片), cv2.COLOR_RGB2BGR)

        # 读取模板图像
        匹配目标图片 = cv2.imread(匹配路径)
        if 匹配目标图片 is None:
            print(f"[×] 无法读取模板图片：{匹配路径}")
            return None

        # 模板匹配
        匹配结果 = cv2.matchTemplate(局部截图图片, 匹配目标图片, cv2.TM_CCOEFF_NORMED)
        _, 最大匹配值, _, 最佳位置 = cv2.minMaxLoc(匹配结果)

        if 最大匹配值 >= 匹配精确度:
            高度, 宽度 = 匹配目标图片.shape[:2]
            中心x = 左上角x + 最佳位置[0] + 宽度 // 2
            中心y = 左上角y + 最佳位置[1] + 高度 // 2
            print(f"[✓] 匹配成功，匹配度：{最大匹配值:.3f}，坐标：({中心x}, {中心y})")
            return (中心x, 中心y)

        print(f"[×] 未匹配成功（匹配度：{最大匹配值:.3f}），{循环间隔}s 后重试...")
        time.sleep(循环间隔)
