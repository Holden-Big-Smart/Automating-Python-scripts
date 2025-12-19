# 文件名：截图工具.py

import pyautogui
import cv2
import numpy as np

def 截图匹配(左上角x, 左上角y, 宽, 高, 匹配路径, 匹配精确度=0.85):
    """
    在指定屏幕区域中查找图像并返回中心坐标

    :param 左上角x: 截图区域左上角 X 坐标
    :param 左上角y: 截图区域左上角 Y 坐标
    :param 宽: 区域宽度
    :param 高: 区域高度
    :param 匹配路径: 模板图像路径
    :param 匹配精确度: 最小匹配相似度（默认0.85）
    :return: (x, y) 匹配成功时返回中心坐标，否则返回 None
    """
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

    # 判断是否匹配成功
    if 最大匹配值 >= 匹配精确度:
        高度, 宽度 = 匹配目标图片.shape[:2]
        中心x = 左上角x + 最佳位置[0] + 宽度 // 2
        中心y = 左上角y + 最佳位置[1] + 高度 // 2
        return (中心x, 中心y)

    # 未匹配成功
    return None
