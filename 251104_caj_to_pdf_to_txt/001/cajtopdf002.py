import os
import win32com.client
from win32com.client import Dispatch

# -------------------------- 配置参数 --------------------------
CAJ_FOLDER = r"H:\all_in_desktop\桌面\课题\2021-09\新时代大学生劳动教育\素材"
PDF_SAVE_FOLDER = r"H:\all_in_desktop\桌面\课题\2021-09\新时代大学生劳动教育\素材"
PRINTER_NAME = "Microsoft Print to PDF"  # 虚拟PDF打印机名称
# --------------------------------------------------------------

os.makedirs(PDF_SAVE_FOLDER, exist_ok=True)
caj_files = [f for f in os.listdir(CAJ_FOLDER) if f.endswith((".caj", ".nh"))]

if not caj_files:
    print("未找到CAJ文件！")
    exit()

# 初始化CAJViewer COM对象
try:
    caj_app = Dispatch("CAJViewer.Application")
    caj_app.Visible = False  # 后台运行（不显示窗口）
except Exception as e:
    print(f"初始化CAJViewer失败：{str(e)}")
    print("请确保已安装CAJViewer 7.3及以上版本！")
    exit()

print(f"找到 {len(caj_files)} 个CAJ文件，开始转换...")

for idx, filename in enumerate(caj_files, 1):
    caj_path = os.path.join(CAJ_FOLDER, filename)
    pdf_filename = os.path.splitext(filename)[0] + ".pdf"
    pdf_path = os.path.join(PDF_SAVE_FOLDER, pdf_filename)

    if os.path.exists(pdf_path):
        print(f"[{idx}/{len(caj_files)}] {filename} → 已存在，跳过")
        continue

    try:
        # 打开CAJ文件
        doc = caj_app.Documents.Open(caj_path)
        time.sleep(1)

        # 配置打印参数（指定打印机、保存路径）
        doc.PrintOut(
            PrinterName=PRINTER_NAME,
            PrintToFile=True,
            OutputFileName=pdf_path,
            Silent=True  # 静默打印（不弹出对话框）
        )

        # 等待打印完成（根据文件大小调整）
        time.sleep(2)

        # 关闭CAJ文件
        doc.Close()
        print(f"[{idx}/{len(caj_files)}] {filename} → 转换成功")

    except Exception as e:
        print(f"[{idx}/{len(caj_files)}] {filename} → 转换失败：{str(e)}")
        try:
            doc.Close()
        except:
            pass

# 退出CAJViewer
caj_app.Quit()
print("批量转换完成！")