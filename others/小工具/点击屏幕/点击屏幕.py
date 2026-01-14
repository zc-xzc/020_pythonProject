import pyautogui
import time
import threading
from pynput import keyboard

# 线程安全的停止信号（替代全局布尔值）
stop_event = threading.Event()
click_thread = None  # 点击线程对象
key_listener = None  # 键盘监听对象


def get_mouse_position():
    """获取目标点击坐标：5秒内将鼠标移到目标位置，自动记录"""
    print("=" * 40)
    print("请在5秒内将鼠标移动到需要自动点击的位置！")
    print("=" * 40)
    time.sleep(5)
    x, y = pyautogui.position()
    print(f"✅ 已记录点击坐标：X={x}, Y={y}")
    return x, y


def auto_click(x, y, interval):
    """自动点击核心逻辑：线程安全，支持即时停止"""
    print(f"\n🚀 自动点击已启动！")
    print(f"👉 点击间隔：{interval}秒 | 停止方式：按 ESC 键")
    while not stop_event.is_set():  # 检测停止信号
        pyautogui.click(x, y)  # 执行点击
        # 用wait替代sleep，支持「即时停止」（sleep期间无法响应停止）
        if stop_event.wait(interval):
            break


def on_key_press(key):
    """键盘监听回调：按ESC键触发停止"""
    try:
        if key == keyboard.Key.esc:
            stop_clicking()
            return False  # 停止键盘监听
    except Exception as e:
        print(f"⚠️ 键盘监听异常：{e}")


def start_clicking(x, y, interval):
    """启动点击线程+键盘监听"""
    global click_thread, key_listener
    stop_event.clear()  # 重置停止信号（确保每次启动都是初始状态）

    # 启动点击线程（设为守护线程，主程序退出时自动终止）
    click_thread = threading.Thread(target=auto_click, args=(x, y, interval))
    click_thread.daemon = True
    click_thread.start()

    # 启动键盘监听（非阻塞，后台运行）
    key_listener = keyboard.Listener(on_press=on_key_press)
    key_listener.start()


def stop_clicking():
    """优雅停止：清理所有线程和资源"""
    if stop_event.is_set():
        return  # 避免重复停止
    print("\n🛑 正在停止自动点击...")
    stop_event.set()  # 触发停止信号

    # 等待点击线程退出（最多等1秒，防止卡死）
    if click_thread and click_thread.is_alive():
        click_thread.join(timeout=1.0)

    # 停止键盘监听
    if key_listener and key_listener.is_alive():
        key_listener.stop()
        key_listener.join(timeout=0.5)

    print("✅ 自动点击已完全停止！")


if __name__ == "__main__":
    try:
        # 1. 获取点击坐标
        target_x, target_y = get_mouse_position()

        # 2. 输入点击间隔（校验合法性）
        while True:
            try:
                interval_input = input("\n请输入点击间隔（秒，例如0.1）：")
                click_interval = float(interval_input)
                if click_interval <= 0:
                    print("❌ 间隔时间必须大于0！请重新输入")
                    continue
                break
            except ValueError:
                print("❌ 输入无效！请输入数字（如0.1、1、2）")

        # 3. 启动自动点击
        start_clicking(target_x, target_y, click_interval)

        # 主线程等待点击线程结束（防止程序直接退出）
        click_thread.join()

    except KeyboardInterrupt:
        stop_clicking()
        print("\n⚠️ 程序被Ctrl+C手动中断")
    except Exception as e:
        stop_clicking()
        print(f"\n❌ 程序运行出错：{str(e)}")