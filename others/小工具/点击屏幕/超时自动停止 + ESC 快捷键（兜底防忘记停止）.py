import pyautogui
import time
import threading
from pynput import keyboard

# 线程安全停止信号
stop_event = threading.Event()
click_thread = None
key_listener = None


def get_mouse_position():
    """获取目标点击坐标"""
    print("=" * 40)
    print("请在5秒内将鼠标移动到需要自动点击的位置！")
    print("=" * 40)
    time.sleep(5)
    x, y = pyautogui.position()
    print(f"✅ 已记录点击坐标：X={x}, Y={y}")
    return x, y


def auto_click(x, y, interval, max_duration):
    """自动点击核心逻辑（支持超时停止）
    :param max_duration: 最大运行时长（秒），None表示无限制
    """
    start_time = time.time()  # 记录启动时间
    print(f"\n🚀 自动点击已启动！")
    print(f"👉 点击间隔：{interval}秒 | 停止方式：ESC键 | 超时自动停止：{max_duration}秒")

    while not stop_event.is_set():
        # 检查是否达到最大运行时长
        if max_duration and (time.time() - start_time) >= max_duration:
            print(f"\n⏰ 达到最大运行时长（{max_duration}秒），自动停止！")
            stop_event.set()
            break

        pyautogui.click(x, y)
        if stop_event.wait(interval):
            break


def on_key_press(key):
    """ESC键停止回调"""
    try:
        if key == keyboard.Key.esc:
            stop_clicking()
            return False
    except Exception as e:
        print(f"⚠️ 键盘监听异常：{e}")


def start_clicking(x, y, interval, max_duration):
    """启动点击+键盘监听"""
    global click_thread, key_listener
    stop_event.clear()

    # 启动点击线程
    click_thread = threading.Thread(target=auto_click, args=(x, y, interval, max_duration))
    click_thread.daemon = True
    click_thread.start()

    # 启动键盘监听
    key_listener = keyboard.Listener(on_press=on_key_press)
    key_listener.start()


def stop_clicking():
    """优雅停止"""
    if stop_event.is_set():
        return
    print("\n🛑 正在停止自动点击...")
    stop_event.set()

    if click_thread and click_thread.is_alive():
        click_thread.join(timeout=1.0)
    if key_listener and key_listener.is_alive():
        key_listener.stop()
        key_listener.join(timeout=0.5)

    print("✅ 自动点击已完全停止！")


if __name__ == "__main__":
    try:
        # 1. 获取点击坐标
        target_x, target_y = get_mouse_position()

        # 2. 输入点击间隔
        while True:
            try:
                interval_input = input("\n请输入点击间隔（秒，例如0.1）：")
                click_interval = float(interval_input)
                if click_interval <= 0:
                    print("❌ 间隔时间必须大于0！")
                    continue
                break
            except ValueError:
                print("❌ 请输入有效数字！")

        # 3. 输入最大运行时长（超时自动停止）
        while True:
            try:
                duration_input = input("请输入最大运行时长（秒，输入0表示无限制）：")
                max_duration = float(duration_input)
                if max_duration < 0:
                    print("❌ 时长不能为负数！")
                    continue
                max_duration = None if max_duration == 0 else max_duration
                break
            except ValueError:
                print("❌ 请输入有效数字！")

        # 4. 启动自动点击
        start_clicking(target_x, target_y, click_interval, max_duration)

        # 主线程等待
        click_thread.join()

    except KeyboardInterrupt:
        stop_clicking()
        print("\n⚠️ 程序被Ctrl+C中断")
    except Exception as e:
        stop_clicking()
        print(f"\n❌ 程序出错：{str(e)}")