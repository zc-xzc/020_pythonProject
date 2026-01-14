import csv
import os
import re
import tkinter as tk
from tkinter import filedialog
from typing import Dict, List, Set

class BatchCsvToBibMerger:
    """批量处理目录/子目录CSV→Bib转换，并合并去重"""
    def __init__(self):
        # 核心配置：CSV列名→BibTeX字段映射（已适配你的SearchResults.csv）
        self.field_mapping = {
            'title': 'title',                # 文献标题
            'authors': 'author',             # 作者
            'year': 'year',                  # 年份
            'journal': 'journal',            # 期刊
            'book title': 'booktitle',       # 会议集/图书名
            'publisher': 'publisher',        # 出版社
            'doi': 'doi',                    # DOI
            'document type': 'entrytype'     # 文献类型（自动适配）
        }
        # 存储所有唯一Bib条目（Key: 条目Key，Value: 条目内容）
        self.all_bib_entries: Dict[str, str] = {}
        # 记录扫描到的CSV文件
        self.found_csv_files: List[str] = []

    def _sanitize_text(self, text: str) -> str:
        """清洗文本：处理特殊字符、多余空格"""
        if not text or str(text).strip().lower() in ['nan', 'none']:
            return ''
        clean_text = re.sub(r'\s+', ' ', str(text).strip())
        # 转义BibTeX特殊字符（{ } \ # $ & ~ _ ^ %）
        special_chars = {'{': '\\{', '}': '\\}', '\\': '\\\\', '#': '\\#',
                         '$': '\\$', '&': '\\&', '~': '\\~', '_': '\\_',
                         '^': '\\^', '%': '\\%'}
        for char, escaped in special_chars.items():
            clean_text = clean_text.replace(char, escaped)
        return clean_text

    def _get_bib_entry_type(self, doc_type: str) -> str:
        """根据CSV的Document Type自动匹配BibTeX条目类型"""
        doc_type = doc_type.lower().strip()
        if 'article' in doc_type or 'journal' in doc_type:
            return 'article'
        elif 'book' in doc_type:
            return 'book'
        elif 'proceeding' in doc_type or 'conference' in doc_type:
            return 'inproceedings'
        elif 'thesis' in doc_type:
            return 'mastersthesis'
        else:
            return 'misc'  # 默认类型

    def _generate_unique_key(self, row: Dict[str, str]) -> str:
        """生成唯一Bib条目Key：作者首字母+年份+标题首字母（去重）"""
        # 提取作者（取第一个作者的姓氏）
        author = self._sanitize_text(row.get('authors', 'Unknown'))
        author_key = author.split(',')[0].split(' ')[-1] if author else 'Unknown'
        # 提取年份（取前4位，无则用0000）
        year = self._sanitize_text(row.get('year', '0000'))[:4]
        # 提取标题首字母（前3个非空单词）
        title = self._sanitize_text(row.get('title', 'NoTitle'))
        title_key = ''.join([w[0].upper() for w in title.split()[:3] if w])
        # 生成基础Key并去重（重复则加数字后缀）
        base_key = f"{author_key}{year}{title_key}"
        final_key = base_key
        counter = 1
        while final_key in self.all_bib_entries:
            final_key = f"{base_key}{counter}"
            counter += 1
        return final_key

    def _csv_row_to_bib(self, row: Dict[str, str]) -> str:
        """单条CSV记录转Bib条目字符串"""
        # 1. 确定条目类型
        doc_type = self._sanitize_text(row.get('document type', ''))
        entry_type = self._get_bib_entry_type(doc_type)
        # 2. 生成唯一Key
        entry_key = self._generate_unique_key(row)
        # 3. 构建Bib字段（只保留非空字段）
        bib_fields = []
        for csv_col, bib_field in self.field_mapping.items():
            if bib_field == 'entrytype':  # 跳过类型字段（已单独处理）
                continue
            value = self._sanitize_text(row.get(csv_col, ''))
            if value:
                bib_fields.append(f"    {bib_field} = {{{value}}}")
        # 4. 拼接完整条目
        fields_str = ',\n'.join(bib_fields)
        return f"@{entry_type}{{{entry_key},\n{fields_str}\n}}"

    def _scan_csv_files(self, root_dir: str) -> None:
        """递归扫描目录及子目录下的所有CSV文件"""
        self.found_csv_files.clear()
        for dirpath, _, filenames in os.walk(root_dir):
            for filename in filenames:
                if filename.lower().endswith('.csv'):
                    csv_path = os.path.join(dirpath, filename)
                    self.found_csv_files.append(csv_path)
        print(f"🔍 扫描完成：在'{root_dir}'及子目录中找到 {len(self.found_csv_files)} 个CSV文件")

    def _convert_single_csv(self, csv_path: str) -> int:
        """转换单个CSV文件为Bib条目，返回成功转换数量"""
        success_count = 0
        # 尝试多种编码读取CSV（避免乱码）
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1', 'utf-8-sig']
        for encoding in encodings:
            try:
                with open(csv_path, 'r', encoding=encoding, newline='') as f:
                    reader = csv.DictReader(f)
                    csv_rows = list(reader)
                # 转换每一行
                for row in csv_rows:
                    try:
                        bib_entry = self._csv_row_to_bib(row)
                        entry_key = self._generate_unique_key(row)
                        self.all_bib_entries[entry_key] = bib_entry
                        success_count += 1
                    except Exception as e:
                        print(f"⚠️  跳过CSV'{csv_path}'中的无效行：{str(e)}")
                print(f"✅ 转换完成：'{os.path.basename(csv_path)}'（成功{success_count}条，编码：{encoding}）")
                return success_count
            except Exception as e:
                continue
        print(f"❌ 读取失败：'{os.path.basename(csv_path)}'（所有编码尝试失败）")
        return 0

    def batch_convert_and_merge(self, root_dir: str, output_bib_path: str) -> None:
        """批量转换+合并主逻辑"""
        try:
            # 1. 扫描所有CSV
            self._scan_csv_files(root_dir)
            if not self.found_csv_files:
                print("⚠️  未找到任何CSV文件，程序退出")
                return
            # 2. 批量转换每个CSV
            total_success = 0
            for csv_path in self.found_csv_files:
                total_success += self._convert_single_csv(csv_path)
            # 3. 合并并写入最终Bib文件
            if self.all_bib_entries:
                # 按Key排序，使输出更整齐
                sorted_bib_entries = sorted(self.all_bib_entries.values(), key=lambda x: x.split('{')[1].split(',')[0])
                with open(output_bib_path, 'w', encoding='utf-8') as f:
                    f.write('\n\n'.join(sorted_bib_entries))
                print(f"\n🎉 全部完成！")
                print(f"📊 统计：共扫描{len(self.found_csv_files)}个CSV，成功转换{total_success}条记录，合并后{len(self.all_bib_entries)}个唯一条目")
                print(f"📁 最终文件：{output_bib_path}")
            else:
                print("⚠️  未生成任何Bib条目（可能所有CSV无有效数据）")
        except Exception as e:
            print(f"❌ 程序出错：{str(e)}")

def select_root_directory():
    """可视化选择CSV所在的根目录（会扫描子目录）"""
    root = tk.Tk()
    root.withdraw()
    dir_path = filedialog.askdirectory(
        title="选择CSV文件所在的根目录（会自动扫描所有子目录）",
        mustexist=True
    )
    root.destroy()
    return dir_path

def select_output_bib_file():
    """可视化选择最终合并Bib文件的保存位置"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        title="选择合并后Bib文件的保存位置",
        defaultextension=".bib",
        filetypes=[("BibTeX文件", "*.bib"), ("所有文件", "*.*")],
        initialfile="merged_all_csv_to_bib.bib"
    )
    root.destroy()
    return file_path

def main():
    print("📌 CSV批量转BibTeX并合并工具（支持目录+子目录）")
    # 1. 选择CSV根目录
    root_dir = select_root_directory()
    if not root_dir:
        print("⚠️  未选择目录，程序退出")
        return
    # 2. 选择输出Bib文件路径
    output_bib_path = select_output_bib_file()
    if not output_bib_path:
        print("⚠️  未选择输出文件，程序退出")
        return
    # 3. 执行批量转换+合并
    converter = BatchCsvToBibMerger()
    converter.batch_convert_and_merge(root_dir, output_bib_path)

if __name__ == "__main__":
    main()