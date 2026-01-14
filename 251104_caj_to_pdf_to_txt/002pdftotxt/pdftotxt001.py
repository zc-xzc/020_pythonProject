import os
import sys
import PyPDF2
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# 手动指定依赖路径（必须与你的安装路径一致！）
os.environ["POPPLER_PATH"] = r'C:\Program Files\poppler-25.07.0\Library\bin'
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def main():
    root = tk.Tk()
    root.title("极简PDF转TXT")
    root.geometry("600x400")

    # 日志区域
    log_text = scrolledtext.ScrolledText(root, state=tk.DISABLED)
    log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def log(msg):
        """添加日志（带时间戳）"""
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, f"[日志] {msg}\n")
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)
        root.update()  # 实时刷新日志

    def select_pdf():
        """选择单个PDF文件（简化为单选，避免多文件问题）"""
        path = filedialog.askopenfilename(
            title="选择单个PDF文件",
            filetypes=[("PDF文件", "*.pdf")]
        )
        if path:
            log(f"已选择PDF: {path}")
            return path
        return None

    def select_output_dir():
        """选择输出目录（默认桌面，简化路径）"""
        path = filedialog.askdirectory(title="选择输出目录")
        if not path:
            # 默认为桌面，避免路径过长
            path = os.path.join(os.path.expanduser("~"), "Desktop", "PDF转TXT输出")
            os.makedirs(path, exist_ok=True)
            log(f"未选择目录，默认输出到: {path}")
        else:
            log(f"已选择输出目录: {path}")
        return path

    def convert_pdf_to_txt():
        """核心转换逻辑（单文件、纯文本模式，去除所有复杂配置）"""
        # 1. 选择文件和目录
        pdf_path = select_pdf()
        if not pdf_path:
            messagebox.showerror("错误", "请选择PDF文件")
            return
        output_dir = select_output_dir()

        # 2. 生成输出文件路径（极简命名，避免非法字符）
        pdf_name = os.path.basename(pdf_path).replace(".pdf", "")
        # 只保留字母、数字、中文和下划线，其他替换为下划线
        valid_name = "".join(c if c.isalnum() or c in "一二三四五六七八九十百千万亿_ " else "_" for c in pdf_name)
        output_path = os.path.join(output_dir, f"{valid_name}_输出.txt")
        log(f"输出路径: {output_path}")

        try:
            # 3. 读取PDF（纯文本模式，最稳定）
            log("开始读取PDF...")
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                # 检查加密
                if reader.is_encrypted:
                    log("错误：PDF文件已加密，无法转换")
                    messagebox.showerror("错误", "PDF文件已加密，请先解密")
                    return
                total_pages = len(reader.pages)
                log(f"PDF总页数: {total_pages}")

                # 4. 提取所有页面文本（不限制页码，避免页码错误）
                log("开始提取文本...")
                all_text = []
                for i in range(total_pages):
                    page_text = reader.pages[i].extract_text()
                    if page_text:
                        all_text.append(f"=== 第{i+1}页 ===\n{page_text}\n")
                        log(f"已提取第{i+1}页文本")
                    else:
                        all_text.append(f"=== 第{i+1}页 ===\n【无可提取文本】\n")
                        log(f"第{i+1}页无文本")

                # 5. 强制写入文件（UTF-8编码，确保兼容性）
                log(f"开始写入文件: {output_path}")
                with open(output_path, 'w', encoding='utf-8', errors='ignore') as f:
                    f.write("".join(all_text))

                # 6. 验证文件是否生成
                if os.path.exists(output_path):
                    file_size = os.path.getsize(output_path)
                    if file_size > 0:
                        log(f"✅ 转换成功！文件大小: {file_size} 字节")
                        log(f"📁 文件位置: {output_path}")
                        messagebox.showinfo("成功", f"转换完成！\n文件已保存至:\n{output_path}")
                        # 自动打开输出目录（方便用户查看）
                        os.startfile(output_dir)
                    else:
                        log("❌ 转换失败：生成的文件为空")
                        messagebox.showerror("失败", "生成的TXT文件为空，可能PDF无文本内容")
                else:
                    log("❌ 转换失败：文件未生成")
                    messagebox.showerror("失败", "文件未生成，请检查输出目录权限")

        except Exception as e:
            log(f"❌ 转换异常: {str(e)}")
            log(f"异常详情: {sys.exc_info()[0]}")
            messagebox.showerror("异常", f"转换出错：{str(e)}\n请查看日志详情")

    # 转换按钮（突出显示）
    convert_btn = tk.Button(
        root,
        text="开始转换（单选PDF+默认桌面输出）",
        command=convert_pdf_to_txt,
        bg="#4CAF50",
        fg="white",
        font=("Arial", 12)
    )
    convert_btn.pack(pady=10)

    log("程序启动成功！点击按钮开始转换（仅支持纯文本PDF）")
    root.mainloop()

if __name__ == "__main__":
    # 前置检查：确保依赖路径存在
    if not os.path.exists(os.environ["POPPLER_PATH"]):
        messagebox.showerror("错误", f"poppler路径不存在：{os.environ['POPPLER_PATH']}\n请修改代码中的路径配置")
        sys.exit(1)
    if not os.path.exists(pytesseract.pytesseract.tesseract_cmd):
        messagebox.showerror("错误", f"Tesseract路径不存在：{pytesseract.pytesseract.tesseract_cmd}\n请修改代码中的路径配置")
        sys.exit(1)
    main()