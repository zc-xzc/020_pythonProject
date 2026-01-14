import os
import sys
import PyPDF2
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import time

# 手动指定依赖路径（必须与安装路径一致！）
POPPLER_PATH = r'C:\Program Files\poppler-25.07.0\Library\bin'
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
os.environ["POPPLER_PATH"] = POPPLER_PATH
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


class BatchPDFToTxtConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("批量PDF转TXT工具（支持大量文件）")
        self.root.geometry("800x600")

        # 状态变量
        self.pdf_files = []  # 批量PDF文件列表
        self.output_dir = ""  # 输出目录
        self.is_running = False  # 转换状态（避免重复运行）
        self.success_count = 0  # 成功数量
        self.fail_count = 0  # 失败数量
        self.total_count = 0  # 总文件数
        self.current_index = 0  # 当前处理索引

        # 创建UI
        self._create_ui()
        # 前置检查依赖路径
        self._check_dependencies()

    def _create_ui(self):
        """创建批量转换界面"""
        # 1. 文件选择区（支持多选）
        file_frame = ttk.LabelFrame(self.root, text="批量选择PDF文件", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(file_frame, text="添加文件（可多选）", command=self._add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="清空列表", command=self._clear_files).pack(side=tk.LEFT, padx=5)
        self.file_count_label = ttk.Label(file_frame, text="已选择：0 个文件")
        self.file_count_label.pack(side=tk.LEFT, padx=5)

        # 2. 输出目录区
        dir_frame = ttk.LabelFrame(self.root, text="输出设置", padding=10)
        dir_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(dir_frame, text="选择输出目录", command=self._select_output_dir).pack(side=tk.LEFT, padx=5)
        self.dir_label = ttk.Label(dir_frame, text="未选择目录（默认：桌面/PDF批量输出）")
        self.dir_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 3. 转换模式区（纯文本优先，OCR可选）
        mode_frame = ttk.LabelFrame(self.root, text="转换模式", padding=10)
        mode_frame.pack(fill=tk.X, padx=10, pady=5)

        self.mode_var = tk.StringVar(value="text")
        ttk.Radiobutton(mode_frame, text="纯文本模式（快，支持可复制文字PDF）", variable=self.mode_var, value="text").pack(
            side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="OCR模式（慢，支持扫描件PDF）", variable=self.mode_var, value="ocr").pack(
            side=tk.LEFT, padx=10)

        # 4. 进度显示区（批量必备）
        progress_frame = ttk.LabelFrame(self.root, text="转换进度", padding=10)
        progress_frame.pack(fill=tk.X, padx=10, pady=5)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=3)

        self.progress_label = ttk.Label(progress_frame, text="进度：0/0 个文件（未开始）")
        self.progress_label.pack(side=tk.LEFT, padx=5)

        # 5. 日志区（详细记录每个文件状态）
        log_frame = ttk.LabelFrame(self.root, text="转换日志（批量处理详情）", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, state=tk.DISABLED, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 6. 操作按钮区
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill=tk.X, padx=10)

        self.start_btn = ttk.Button(btn_frame, text="开始批量转换", command=self._start_batch_conversion,
                                    state=tk.NORMAL)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(btn_frame, text="停止转换", command=self._stop_conversion, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="打开输出目录", command=self._open_output_dir).pack(side=tk.RIGHT, padx=5)

    def _log(self, msg, is_error=False):
        """添加日志（错误信息标红）"""
        self.log_text.config(state=tk.NORMAL)
        if is_error:
            self.log_text.insert(tk.END, f"[错误] {time.strftime('%H:%M:%S')} - {msg}\n", "error")
        else:
            self.log_text.insert(tk.END, f"[信息] {time.strftime('%H:%M:%S')} - {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()  # 实时刷新UI

    def _add_files(self):
        """批量添加PDF文件（支持多选）"""
        files = filedialog.askopenfilenames(
            title="选择批量PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if files:
            new_files = [f for f in files if f not in self.pdf_files]  # 去重
            self.pdf_files.extend(new_files)
            self.total_count = len(self.pdf_files)
            self.file_count_label.config(text=f"已选择：{self.total_count} 个文件")
            self._log(f"添加 {len(new_files)} 个文件，累计 {self.total_count} 个")

    def _clear_files(self):
        """清空文件列表"""
        self.pdf_files = []
        self.total_count = 0
        self.current_index = 0
        self.success_count = 0
        self.fail_count = 0
        self.file_count_label.config(text="已选择：0 个文件")
        self.progress_var.set(0)
        self.progress_label.config(text="进度：0/0 个文件（未开始）")
        self._log("已清空文件列表")

    def _select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择批量输出目录")
        if dir_path:
            self.output_dir = dir_path
            self.dir_label.config(text=f"输出目录：{dir_path}")
            self._log(f"已选择输出目录：{dir_path}")
        else:
            self._set_default_output_dir()

    def _set_default_output_dir(self):
        """设置默认输出目录（桌面/批量输出）"""
        self.output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "PDF批量转TXT输出")
        os.makedirs(self.output_dir, exist_ok=True)
        self.dir_label.config(text=f"默认输出目录：{self.output_dir}")
        self._log(f"使用默认输出目录：{self.output_dir}")

    def _open_output_dir(self):
        """打开输出目录"""
        if os.path.exists(self.output_dir):
            os.startfile(self.output_dir)
        else:
            messagebox.showwarning("警告", "输出目录不存在！")

    def _check_dependencies(self):
        """检查依赖路径是否有效"""
        if not os.path.exists(POPPLER_PATH):
            messagebox.showerror("错误", f"poppler路径不存在：{POPPLER_PATH}\n请修改代码中的路径配置")
            self.root.quit()
        if not os.path.exists(TESSERACT_PATH):
            messagebox.showerror("错误", f"Tesseract路径不存在：{TESSERACT_PATH}\n请修改代码中的路径配置")
            self.root.quit()
        self._set_default_output_dir()  # 默认输出目录
        self._log("依赖检查通过，程序就绪")

    def _clean_filename(self, filename):
        """清洗文件名（避免非法字符导致写入失败）"""
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename[:50]  # 限制文件名长度（避免路径过长）

    def _convert_single_pdf(self, pdf_path):
        """转换单个PDF文件（核心逻辑，确保输出）"""
        pdf_name = os.path.basename(pdf_path)
        clean_name = self._clean_filename(os.path.splitext(pdf_name)[0])
        mode = self.mode_var.get()

        # 生成输出路径（区分模式）
        suffix = "_ocr.txt" if mode == "ocr" else ".txt"
        output_path = os.path.join(self.output_dir, f"{clean_name}{suffix}")

        try:
            # 检查PDF文件是否存在
            if not os.path.exists(pdf_path):
                self._log(f"{pdf_name} - 文件不存在", is_error=True)
                return False

            # 纯文本模式转换（快，优先推荐）
            if mode == "text":
                with open(pdf_path, 'rb') as f:
                    reader = PyPDF2.PdfReader(f)
                    # 处理加密PDF
                    if reader.is_encrypted:
                        self._log(f"{pdf_name} - PDF已加密，跳过", is_error=True)
                        return False
                    total_pages = len(reader.pages)
                    self._log(f"{pdf_name} - 开始转换（纯文本模式，共{total_pages}页）")

                    # 提取所有页面文本
                    all_text = []
                    for i in range(total_pages):
                        page_text = reader.pages[i].extract_text() or f"【第{i + 1}页无文本】"
                        all_text.append(f"=== 第{i + 1}页 ===\n{page_text}\n")

            # OCR模式转换（扫描件专用）
            else:
                self._log(f"{pdf_name} - 开始转换（OCR模式，较慢，请耐心等待）")
                # PDF转图片
                images = convert_from_path(
                    pdf_path,
                    dpi=200,  # 降低dpi加快速度（批量场景优先效率）
                    fmt="png",
                    thread_count=2,
                    timeout=30
                )
                if not images:
                    self._log(f"{pdf_name} - OCR模式未提取到页面", is_error=True)
                    return False

                # OCR识别文本
                all_text = []
                for i, img in enumerate(images):
                    page_text = pytesseract.image_to_string(img, lang="chi_sim+eng") or f"【第{i + 1}页OCR识别失败】"
                    all_text.append(f"=== 第{i + 1}页 ===\n{page_text.strip()}\n")

            # 强制写入文件（确保生成）
            with open(output_path, 'w', encoding='utf-8', errors='ignore') as f:
                f.write("".join(all_text))

            # 验证文件是否生成成功
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                file_size = os.path.getsize(output_path) / 1024  # KB
                self._log(f"{pdf_name} - 转换成功！文件大小：{file_size:.1f}KB，路径：{output_path}")
                return True
            else:
                self._log(f"{pdf_name} - 转换失败：生成文件为空或未创建", is_error=True)
                return False

        except Exception as e:
            self._log(f"{pdf_name} - 转换异常：{str(e)}", is_error=True)
            return False

    def _batch_conversion_thread(self):
        """批量转换线程（避免UI阻塞）"""
        self.is_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.success_count = 0
        self.fail_count = 0

        self._log(f"开始批量转换，共 {self.total_count} 个文件，模式：{self.mode_var.get()}")

        for idx, pdf_path in enumerate(self.pdf_files):
            if not self.is_running:
                self._log("批量转换已停止")
                break

            self.current_index = idx + 1
            self._log(f"\n===== 处理第 {self.current_index}/{self.total_count} 个文件 =====")

            # 转换单个文件
            if self._convert_single_pdf(pdf_path):
                self.success_count += 1
            else:
                self.fail_count += 1

            # 更新进度
            progress = (self.current_index / self.total_count) * 100
            self.progress_var.set(progress)
            self.progress_label.config(
                text=f"进度：{self.current_index}/{self.total_count} 个文件（成功：{self.success_count}，失败：{self.fail_count}）")

        # 转换完成
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self._log(f"\n===== 批量转换结束 =====")
        self._log(f"总文件数：{self.total_count}，成功：{self.success_count}，失败：{self.fail_count}")

        # 自动打开输出目录
        if self.success_count > 0:
            messagebox.showinfo("完成",
                                f"批量转换结束！\n成功：{self.success_count} 个，失败：{self.fail_count} 个\n已自动打开输出目录")
            self._open_output_dir()
        else:
            messagebox.showwarning("完成", f"批量转换结束！\n所有文件转换失败，请查看日志排查问题")

    def _start_batch_conversion(self):
        """启动批量转换（新线程运行）"""
        if not self.pdf_files:
            messagebox.showerror("错误", "请先添加PDF文件！")
            return

        # 确认开始
        confirm = messagebox.askyesno("确认", f"即将批量转换 {self.total_count} 个文件，是否开始？")
        if confirm:
            threading.Thread(target=self._batch_conversion_thread, daemon=True).start()

    def _stop_conversion(self):
        """停止批量转换"""
        self.is_running = False
        self._log("正在停止批量转换...（当前文件处理完成后停止）")


if __name__ == "__main__":
    root = tk.Tk()
    app = BatchPDFToTxtConverter(root)
    # 配置日志颜色（错误标红）
    app.log_text.tag_configure("error", foreground="red")
    root.mainloop()