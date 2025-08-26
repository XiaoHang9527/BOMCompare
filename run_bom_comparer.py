#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
BOM对比工具启动脚本
"""

import os
import sys
import traceback
from bom_comparer import BOMComparerGUI
import tkinter as tk

def main():
    """主函数"""
    try:
        # 创建主窗口
        root = tk.Tk()

        # 设置应用图标
        try:
            root.iconbitmap("icon.ico")
        except:
            pass  # 图标不存在时忽略

        # 获取屏幕尺寸
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # 设置默认窗口尺寸
        default_width = 1200  # 默认宽度
        default_height = 900  # 默认高度

        # 调整窗口大小，适应小屏幕
        if screen_height <= 900:  # 对于小屏幕，进一步减小窗口高度
            default_height = min(default_height, int(screen_height * 0.85))  # 最多使用屏幕高度的85%
            default_width = min(default_width, int(screen_width * 0.85))  # 最多使用屏幕宽度的85%

        # 计算窗口位置 - 使用固定偏移而不是百分比
        x = (screen_width - default_width) // 2  # 水平居中
        y = 20  # 固定距离屏幕顶部20像素

        # 确保坐标不为负
        x = max(0, x)
        y = max(0, y)

        # 设置初始窗口尺寸和位置
        root.geometry(f"{default_width}x{default_height}+{x}+{y}")

        # 打印调试信息
        print(f"屏幕尺寸: {screen_width}x{screen_height}")
        print(f"窗口尺寸: {default_width}x{default_height}")
        print(f"窗口位置: +{x}+{y}")

        # 创建应用
        app = BOMComparerGUI(root)

        # 运行应用
        root.mainloop()
    except Exception as e:
        print(f"程序启动出错: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main()