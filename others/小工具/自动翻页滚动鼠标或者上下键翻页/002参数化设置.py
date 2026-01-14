import time
import threading
import pyautogui
from pynput import keyboard

# ===================== 默认参数配置 =====================
DEFAULT_INTERVAL = 2  # 自动翻页间隔（秒，两种模式通用）
DEFAULT_SCROLL_PIXEL = 300  # 自动滚轮翻页的像素（仅滚轮模式用）
DEFAULT_PAUSE_TIME = 5  # 鼠标定位+窗口激活准备时间（秒，通用）

# 全局控制变量
running = True
selected_mode = None  # 选中模式：'scroll'（滚轮） / 'key'（按键）
auto_interval = None  # 自动执行间隔（秒）
scroll_direction = None  # 翻页方向：'up'（向上） / 'down'（向下）
scroll_pixel = None  # 滚轮翻页像素（仅滚轮模式）
pause_time = None  # 鼠标定位+窗口激活时间


# -------------------- 工具函数：输入验证 --------------------
def validate_num_input(input_str, param_name, min_val, default_val):
    """验证数值输入合法性，无效则返回默认值"""
    if not input_str.strip():
        return default_val
    try:
        val = float(input_str)
        if val < min_val:
            print(f"⚠️ {param_name}不能小于{min_val}，使用默认值{default_val}")
            return default_val
        return val
    except ValueError:
        print(f"⚠️ {param_name}输入无效（需为数字），使用默认值{default_val}")
        return default_val


# -------------------- 模式1：自动滚轮翻页（纯自动） --------------------
def config_scroll_mode():
    """配置自动滚轮翻页模式参数"""
    global auto_interval, scroll_direction, scroll_pixel, pause_time

    # 1. 选择翻页方向
    print("\n===== 自动滚轮模式 - 选择翻页方向 =====")
    while True:
        dir_input = input("翻页方向（u=向上 / d=向下，回车默认向下）：").strip().lower()
        if not dir_input or dir_input == 'd':
            scroll_direction = 'down'
            print(f"✅ 翻页方向：向下")
            break
        elif dir_input == 'u':
            scroll_direction = 'up'
            print(f"✅ 翻页方向：向上")
            break
        else:
            print("❌ 输入错误！仅支持 u（向上） / d（向下）")

    # 2. 设置自动执行间隔
    print("\n===== 自动滚轮模式 - 设置执行间隔 =====")
    interval_input = input(f"自动翻页间隔（秒），回车默认{DEFAULT_INTERVAL}秒：")
    auto_interval = validate_num_input(interval_input, "执行间隔", 0.1, DEFAULT_INTERVAL)
    print(f"✅ 自动翻页间隔：{auto_interval}秒")

    # 3. 设置滚轮滚动像素
    print("\n===== 自动滚轮模式 - 设置滚动像素 =====")
    pixel_input = input(f"单次滚轮滚动像素，回车默认{DEFAULT_SCROLL_PIXEL}像素：")
    scroll_pixel = validate_num_input(pixel_input, "滚动像素", 1, DEFAULT_SCROLL_PIXEL)
    print(f"✅ 单次滚动像素：{scroll_pixel}像素")

    # 4. 设置鼠标定位+窗口激活时间
    print("\n===== 自动滚轮模式 - 设置准备时间 =====")
    pause_input = input(f"鼠标定位+窗口激活准备时间（秒），回车默认{DEFAULT_PAUSE_TIME}秒：")
    pause_time = validate_num_input(pause_input, "准备时间", 1, DEFAULT_PAUSE_TIME)
    print(f"✅ 准备时间：{pause_time}秒")


def auto_scroll_task():
    """自动滚轮翻页核心任务（纯自动，按间隔滚动滚轮）"""
    # 适配pyautogui滚轮方向：负数向下，正数向上
    actual_pixel = -scroll_pixel if scroll_direction == 'down' else scroll_pixel
    while running:
        pyautogui.scroll(actual_pixel)
        print(f"📌 自动滚轮{scroll_direction}翻页 {scroll_pixel} 像素（间隔{auto_interval}秒）")
        time.sleep(auto_interval)


# -------------------- 模式2：自动按键翻页（纯自动，优化版） --------------------
def config_key_mode():
    """配置自动按键翻页模式参数"""
    global auto_interval, scroll_direction, pause_time

    # 1. 选择翻页方向（对应自动按↑/↓键）
    print("\n===== 自动按键模式 - 选择翻页方向 =====")
    while True:
        dir_input = input("翻页方向（u=向上/按↑键 / d=向下/按↓键，回车默认向下）：").strip().lower()
        if not dir_input or dir_input == 'd':
            scroll_direction = 'down'
            print(f"✅ 翻页方向：向下（自动按↓键）")
            break
        elif dir_input == 'u':
            scroll_direction = 'up'
            print(f"✅ 翻页方向：向上（自动按↑键）")
            break
        else:
            print("❌ 输入错误！仅支持 u（向上） / d（向下）")

    # 2. 设置自动执行间隔
    print("\n===== 自动按键模式 - 设置执行间隔 =====")
    interval_input = input(f"自动翻页间隔（秒），回车默认{DEFAULT_INTERVAL}秒：")
    auto_interval = validate_num_input(interval_input, "执行间隔", 0.1, DEFAULT_INTERVAL)
    print(f"✅ 自动翻页间隔：{auto_interval}秒")

    # 3. 设置鼠标定位+窗口激活时间
    print("\n===== 自动按键模式 - 设置准备时间 =====")
    pause_input = input(f"鼠标定位+窗口激活准备时间（秒），回车默认{DEFAULT_PAUSE_TIME}秒：")
    pause_time = validate_num_input(pause_input, "准备时间", 1, DEFAULT_PAUSE_TIME)
    print(f"✅ 准备时间：{pause_time}秒")


def auto_key_task():
    """自动按键翻页核心任务（优化版：模拟真实按键按下/释放，解决没反应问题）"""
    # 映射方向到对应按键
    key_map = {'up': 'up', 'down': 'down'}
    target_key = key_map[scroll_direction]
    press_duration = 0.2  # 按键按下时长（模拟真实操作，避免过快）
    while running:
        # 模拟「按下按键→等待→释放按键」，而非瞬间按放
        pyautogui.keyDown(target_key)
        time.sleep(press_duration)
        pyautogui.keyUp(target_key)
        print(f"📌 自动按{target_key.upper()}键翻页（间隔{auto_interval}秒）")
        # 总间隔扣除按键按下时长，保证间隔准确
        time.sleep(auto_interval - press_duration)


# -------------------- 通用退出监听 --------------------
def esc_exit_listen(key):
    """仅监听ESC键退出脚本"""
    global running
    if key == keyboard.Key.esc:
        running = False
        print("\n✅ 脚本已退出")
        return False  # 停止监听


# -------------------- 主流程控制 --------------------
def main():
    global selected_mode, running
    auto_task_thread = None

    # 第一步：选择纯自动翻页的操作方式
    print("===== 纯自动翻页脚本（修复按键无响应） =====")
    while True:
        mode_input = input(
            "请选择自动翻页的操作方式：\n  输入 s → 自动滚轮翻页（脚本滚鼠标）\n  输入 k → 自动按键翻页（脚本按↑/↓键）\n  你的选择：").strip().lower()
        if mode_input == 's':
            selected_mode = 'scroll'
            print(f"\n✅ 已选择【自动滚轮翻页模式】（纯自动，脚本滚动鼠标）")
            break
        elif mode_input == 'k':
            selected_mode = 'key'
            print(f"\n✅ 已选择【自动按键翻页模式】（纯自动，脚本按上下键）")
            break
        else:
            print("❌ 输入错误！仅支持 s（滚轮） / k（按键）\n")

    # 第二步：配置对应模式的参数
    if selected_mode == 'scroll':
        config_scroll_mode()
    else:
        config_key_mode()

    # 第三步：提示鼠标定位+激活窗口（关键修复点）
    print(f"\n===== {selected_mode.upper()}模式 - 即将启动 =====")
    print(f"📢 请在{pause_time}秒内完成以下操作（关键！）：")
    print(f"   1. 将鼠标移到目标操作窗口（如浏览器、PDF、Word）；")
    print(f"   2. 点击目标窗口的空白处，确保窗口处于【激活状态】；")
    if selected_mode == 'scroll':
        print(f"🔧 运行规则：每{auto_interval}秒自动{scroll_direction}滚动{scroll_pixel}像素，按ESC退出")
    else:
        print(f"🔧 运行规则：每{auto_interval}秒自动按{scroll_direction.upper()}键翻页，按ESC退出")

    # 第四步：等待鼠标定位+窗口激活
    time.sleep(pause_time)
    print(f"\n✅ {selected_mode.upper()}模式已启动！全程纯自动执行，按ESC退出")

    # 第五步：启动纯自动翻页任务
    if selected_mode == 'scroll':
        auto_task_thread = threading.Thread(target=auto_scroll_task)
    else:
        auto_task_thread = threading.Thread(target=auto_key_task)
    auto_task_thread.daemon = True  # 守护线程，主程序退出时自动结束
    auto_task_thread.start()

    # 启动ESC键退出监听（阻塞直到退出）
    with keyboard.Listener(on_press=esc_exit_listen) as listener:
        listener.join()

    # 等待自动任务线程结束
    if auto_task_thread and auto_task_thread.is_alive():
        auto_task_thread.join()


if __name__ == "__main__":
    # pyautogui基础配置：避免操作过快/误触退出
    pyautogui.PAUSE = 0.1  # 每次操作后暂停0.1秒，防止过快
    pyautogui.FAILSAFE = False  # 禁用左上角失败安全（避免鼠标移到左上角误退出）

    try:
        main()
    except KeyboardInterrupt:
        running = False
        print("\n⚠️ 脚本被Ctrl+C手动终止")
    except Exception as e:
        running = False
        print(f"\n❌ 脚本运行出错：{e}")