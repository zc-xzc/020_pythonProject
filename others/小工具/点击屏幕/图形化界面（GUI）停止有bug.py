import pyautogui
import time
import threading
import tkinter as tk
from tkinter import messagebox, simpledialog

# 线程安全停止信号
stop_event = threading.Event()
click_thread = None


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
    while not stop_event.is_set():
        pyautogui.click(x, y)
        if stop_event.wait(interval):
            break


def start_clicking():
    """GUI启动按钮回调"""
    global click_thread
    try:
        # 1. 获取点击坐标
        target_x, target_y = get_mouse_position()

        # 2. 获取用户输入的间隔（从GUI输入框）
        interval_input = entry_interval.get().strip()
        if not interval_input:
            messagebox.showerror("错误", "请输入点击间隔！")
            return
        click_interval = float(interval_input)

        if click_interval <= 0:
            messagebox.showerror("错误", "间隔时间必须大于0！")
            return

        # 3. 重置停止信号，启动点击线程
        stop_event.clear()
        click_thread = threading.Thread(target=auto_click, args=(target_x, target_y, click_interval))
        click_thread.daemon = True
        click_thread.start()

        messagebox.showinfo("启动成功",
                            f"自动点击已启动！\n点击间隔：{click_interval}秒\n停止方式：点击「停止点击」按钮")

    except ValueError:
        messagebox.showerror("错误", "请输入有效的数字（如0.1、1）！")
    except Exception as e:
        messagebox.showerror("错误", f"启动失败：{str(e)}")


def stop_clicking():
    """GUI停止按钮回调"""
    if stop_event.is_set():
        messagebox.showinfo("提示", "自动点击已停止！")
        return

    stop_event.set()
    # 等待线程退出
    if click_thread and click_thread.is_alive():
        click_thread.join(timeout=1.0)

    messagebox.showinfo("停止成功", "✅ 自动点击已完全停止！")
    root.quit()  # 关闭GUI窗口


# ------------------- 创建GUI界面 -------------------
if __name__ == "__main__":
    # 主窗口配置
    root = tk.Tk()
    root.title("自动点击器 - 图形化版")
    root.geometry("350x200")  # 窗口大小
    root.resizable(False, False)  # 禁止调整窗口大小

    # 标题标签
    label_title = tk.Label(root, text="自动点击器", font=("Arial", 16, "bold"))
    label_title.pack(pady=10)

    # 间隔输入框
    label_interval = tk.Label(root, text="点击间隔（秒）：", font=("Arial", 12))
    label_interval.pack(pady=5)
    entry_interval = tk.Entry(root, font=("Arial", 12), width=20)
    entry_interval.pack(pady=5)
    entry_interval.insert(0, "0.1")  # 默认值

    # 按钮区域
    frame_btn = tk.Frame(root)
    frame_btn.pack(pady=15)

    btn_start = tk.Button(
        frame_btn, text="启动点击", font=("Arial", 12),
        width=10, bg="#4CAF50", fg="white", command=start_clicking
    )
    btn_start.grid(row=0, column=0, padx=10)

    btn_stop = tk.Button(
        frame_btn, text="停止点击", font=("Arial", 12),
        width=10, bg="#f44336", fg="white", command=stop_clicking
    )
    btn_stop.grid(row=0, column=1, padx=10)

    # 运行GUI主循环
    root.mainloop()