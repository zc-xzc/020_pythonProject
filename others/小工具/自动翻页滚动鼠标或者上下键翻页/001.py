import time
import threading
import pyautogui
from pynput import keyboard

# ===================== 可自定义参数 =====================
SCROLL_INTERVAL = 2  # 自动滚动间隔（秒），可自行修改
SCROLL_AMOUNT = 300  # 每次滚动的距离（像素），正数向下，负数向上
PAUSE_BEFORE_START = 5  # 启动后等待几秒让你移动鼠标（秒）
# ========================================================

# 控制线程运行的标志
running = True
auto_scroll_thread = None


def auto_scroll():
    """自动滚动函数，按设定间隔执行滚动"""
    while running:
        # 滚动鼠标滚轮（第二个参数是垂直滚动距离）
        pyautogui.scroll(-SCROLL_AMOUNT)  # pyautogui中，负数向下滚，正数向上滚
        time.sleep(SCROLL_INTERVAL)


def on_press(key):
    """键盘按键监听函数"""
    global running
    try:
        # 监听上下方向键
        if key == keyboard.Key.down:
            # 向下滚动
            pyautogui.scroll(-SCROLL_AMOUNT)
        elif key == keyboard.Key.up:
            # 向上滚动
            pyautogui.scroll(SCROLL_AMOUNT)
        # 按ESC键退出脚本
        elif key == keyboard.Key.esc:
            running = False
            print("脚本已退出")
            return False  # 停止键盘监听
    except Exception as e:
        print(f"按键监听出错: {e}")


def main():
    global auto_scroll_thread
    print(f"脚本将在{PAUSE_BEFORE_START}秒后启动，请将鼠标移动到需要操作的位置！")
    print("操作说明：")
    print(f"1. 自动模式：每{SCROLL_INTERVAL}秒自动向下滚动{SCROLL_AMOUNT}像素")
    print(f"2. 手动模式：按↑键向上滚动，按↓键向下滚动")
    print("3. 按ESC键退出脚本")

    # 等待指定时间，让你移动鼠标到目标位置
    time.sleep(PAUSE_BEFORE_START)

    # 启动自动滚动线程
    auto_scroll_thread = threading.Thread(target=auto_scroll)
    auto_scroll_thread.daemon = True  # 守护线程，主程序退出时自动结束
    auto_scroll_thread.start()

    # 启动键盘监听
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

    # 等待自动滚动线程结束
    if auto_scroll_thread.is_alive():
        auto_scroll_thread.join()


if __name__ == "__main__":
    # 设置pyautogui的暂停时间（避免操作过快）
    pyautogui.PAUSE = 0.1
    # 禁用pyautogui的失败安全（防止鼠标移到左上角触发退出）
    pyautogui.FAILSAFE = False
    try:
        main()
    except KeyboardInterrupt:
        running = False
        print("\n脚本被手动终止")
    except Exception as e:
        running = False
        print(f"脚本运行出错: {e}")