import os
import time
import pyautogui
import pygetwindow as gw

# -------------------------- 核心配置（必须核对修改！）--------------------------
CAJ_FOLDER = r"H:\all_in_desktop\桌面\课题\2021-09\新时代大学生劳动教育\素材"  # 例如：r"D:\CAJ文件"（把你的CAJ文件都放这里）
PDF_SAVE_FOLDER = r"H:\all_in_desktop\桌面\课题\2021-09\新时代大学生劳动教育\素材"  # 转换后的PDF保存路径（自动创建，不用手动建）
CAJVIEWER_PATH = r"D:\Program Files\TTKN\CAJViewer 8.1\CAJVieweru.exe"  # 核对CAJViewer安装路径
DELAY = 3  # 操作延迟（电脑慢就改5，快就保持3）
# ------------------------------------------------------------------------------

# 确保输出文件夹存在
os.makedirs(PDF_SAVE_FOLDER, exist_ok=True)

# 获取所有CAJ/NH文件
caj_files = [f for f in os.listdir(CAJ_FOLDER) if f.endswith((".caj", ".nh"))]
if not caj_files:
    print("❌ 未在指定文件夹找到CAJ/NH文件！")
    exit()

print(f"✅ 找到 {len(caj_files)} 个CAJ文件，开始转换（运行时不要动鼠标键盘！）...")

for idx, filename in enumerate(caj_files, 1):
    caj_path = os.path.join(CAJ_FOLDER, filename)
    pdf_filename = os.path.splitext(filename)[0] + ".pdf"
    pdf_path = os.path.join(PDF_SAVE_FOLDER, pdf_filename)

    # 跳过已转换的文件
    if os.path.exists(pdf_path):
        print(f"[{idx}/{len(caj_files)}] 🚫 {filename} → 已存在，跳过")
        continue

    try:
        # 1. 用CAJViewer打开文件
        os.startfile(caj_path)
        time.sleep(DELAY * 2)  # 等待CAJViewer加载（大文件可改3倍DELAY）

        # 2. 激活CAJViewer窗口（防止操作没聚焦）
        caj_windows = gw.getWindowsWithTitle("CAJViewer")
        if caj_windows:
            caj_windows[0].activate()
            time.sleep(DELAY)
        else:
            raise Exception("CAJViewer窗口未找到")

        # 3. Ctrl+P打开打印对话框
        pyautogui.hotkey("ctrl", "p")
        time.sleep(DELAY * 1.5)

        # 4. 选择「Microsoft Print to PDF」打印机（关键步骤）
        pyautogui.hotkey("alt", "d")  # 聚焦打印机选择框
        time.sleep(DELAY)
        pyautogui.typewrite("Microsoft Print to PDF")  # 输入打印机名称
        pyautogui.press("enter")
        time.sleep(DELAY)

        # 5. 确认打印（按Enter）
        pyautogui.press("enter")
        time.sleep(DELAY * 2)

        # 6. 输入PDF保存路径并确认
        pyautogui.typewrite(pdf_path)
        pyautogui.press("enter")
        time.sleep(DELAY * 2)  # 等待打印转换（大文件改3-4倍DELAY）

        # 7. 关闭CAJViewer（Alt+F4）
        pyautogui.hotkey("alt", "f4")
        time.sleep(DELAY)

        print(f"[{idx}/{len(caj_files)}] ✅ {filename} → 转换成功")

    except Exception as e:
        print(f"[{idx}/{len(caj_files)}] ❌ {filename} → 转换失败：{str(e)}")
        # 强制关闭可能卡住的窗口
        pyautogui.hotkey("alt", "f4")
        time.sleep(DELAY)

print("\n🎉 批量转换完成！PDF文件保存在：", PDF_SAVE_FOLDER)