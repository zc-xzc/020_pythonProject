import pyautogui
import time
import threading
from pynput import mouse

# 线程安全停止信号
stop_event = threading.Event()
click_thread = None
mouse_listener = None
click_count = 0  # 记录鼠标点击次数（用于检测双击）
last_click_time = 0  # 上一次点击时间（判断双击间隔）
DOUBLE_CLICK_THRESHOLD = 0.5  # 双击判定阈值（秒）


def get_mouse_position():
    """获取目标点击坐标"""
    print("=" * 40)
    print("请在5秒内将鼠标移动到需要自动点击的位置！")
    print("=" * 40)
    time.sleep(5)
    x, y = pyautogui.position()
    print(f"✅ 已记录点击坐标：X={x}, Y={y}")
    return x, y


def auto_click(x, y, interval):
    """自动点击核心逻辑"""
    print(f"\n🚀 自动点击已启动！")
    print(f"👉 点击间隔：{interval}秒 | 停止方式：双击任意位置")
    while not stop_event.is_set():
        pyautogui.click(x, y)
        if stop_event.wait(interval):
            break


def on_mouse_click(x, y, button, pressed):
    """鼠标点击监听：双击左键停止"""
    global click_count, last_click_time
    if button == mouse.Button.left and pressed:  # 只监听左键按下
        current_time = time.time()
        # 判断是否为双击（两次点击间隔<阈值）
        if current_time - last_click_time < DOUBLE_CLICK_THRESHOLD:
            click_count += 1
        else:
            click_count = 1  # 重置计数
        last_click_time = current_time

        # 双击触发停止
        if click_count >= 2:
            stop_clicking()
            return False  # 停止鼠标监听


def start_clicking(x, y, interval):
    """启动点击+鼠标监听"""
    global click_thread, mouse_listener
    stop_event.clear()

    # 启动点击线程
    click_thread = threading.Thread(target=auto_click, args=(x, y, interval))
    click_thread.daemon = True
    click_thread.start()

    # 启动鼠标监听（后台运行）
    mouse_listener = mouse.Listener(on_click=on_mouse_click)
    mouse_listener.start()


def stop_clicking():
    """优雅停止"""
    if stop_event.is_set():
        return
    print("\n🛑 检测到鼠标双击，正在停止自动点击...")
    stop_event.set()

    if click_thread and click_thread.is_alive():
        click_thread.join(timeout=1.0)
    if mouse_listener and mouse_listener.is_alive():
        mouse_listener.stop()
        mouse_listener.join(timeout=0.5)

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

        # 3. 启动自动点击
        start_clicking(target_x, target_y, click_interval)

        # 主线程等待
        click_thread.join()

    except KeyboardInterrupt:
        stop_clicking()
        print("\n⚠️ 程序被Ctrl+C中断")
    except Exception as e:
        stop_clicking()
        print(f"\n❌ 程序出错：{str(e)}")