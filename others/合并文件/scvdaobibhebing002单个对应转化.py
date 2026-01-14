import csv
import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Dict, List, Tuple, Set


class SingleCsvToBib:
    """单CSV→单Bib转换工具（适配Zotero，带文件存在检测）"""

    def __init__(self):
        # 参考Bib模板（字段顺序+必填字段）
        self.ref_templates: Dict[str, Tuple[List[str], Set[str]]] = {}
        # CSV列名→Bib字段映射（适配你的SearchResults.csv）
        self.csv_bib_mapping = {
            'Item Title': 'title',
            'Authors': 'author',
            'Publication Year': 'year',
            'Publication Title': 'journal',
            'Book Series Title': 'booktitle',
            'Journal Volume': 'volume',
            'Journal Issue': 'number',
            'Pages': 'pages',
            'Item DOI': 'doi',
            'URL': 'url',
            'Content Type': 'entrytype',
            'Editor': 'editor'
        }
        # 统计信息（新增：跳过数）
        self.success_count = 0  # 新建成功
        self.skip_count = 0  # 已存在跳过
        self.fail_count = 0  # 转换失败
        self.generated_files = []

    def _parse_reference_file(self, ref_path: str) -> bool:
        """解析参考Bib/TXT文件，提取Zotero兼容模板"""
        try:
            with open(ref_path, 'r', encoding='utf-8') as f:
                content = f.read()

            entry_pattern = re.compile(
                r'@(?P<type>\w+)\s*{\s*(?P<key>[^,]+)\s*,[\s\S]*?^\s*}',
                re.MULTILINE | re.IGNORECASE
            )
            field_pattern = re.compile(r'^\s*(?P<field>\w+)\s*=\s*\{(?P<value>.*?)\}', re.MULTILINE | re.DOTALL)

            for match in entry_pattern.finditer(content):
                entry_type = match.group('type').lower()
                if entry_type in self.ref_templates:
                    continue

                entry_content = match.group(0)
                field_matches = field_pattern.finditer(entry_content)
                fields_order = [m.group('field').lower() for m in field_matches if
                                m.group('field').lower() not in ['key']]
                required_fields = set(fields_order[:5])
                self.ref_templates[entry_type] = (fields_order, required_fields)

            if not self.ref_templates:
                print("⚠️  参考文件解析失败，使用默认Zotero模板")
                self._set_default_template()
                return False
            else:
                print(f"✅ 参考文件解析成功！提取{len(self.ref_templates)}种条目模板")
                return True
        except Exception as e:
            print(f"❌ 参考文件读取错误：{str(e)}，使用默认模板")
            self._set_default_template()
            return False

    def _set_default_template(self):
        """默认Zotero兼容模板"""
        self.ref_templates = {
            'article': (
                ['author', 'title', 'journal', 'volume', 'number', 'pages', 'year', 'issn', 'doi', 'url', 'keywords',
                 'abstract'],
                {'author', 'title', 'journal', 'year', 'doi'}
            ),
            'incollection': (
                ['author', 'title', 'booktitle', 'editor', 'publisher', 'edition', 'address', 'pages', 'year', 'isbn',
                 'doi', 'url', 'keywords', 'abstract'],
                {'author', 'title', 'booktitle', 'publisher', 'year'}
            ),
            'inproceedings': (
                ['author', 'title', 'booktitle', 'publisher', 'pages', 'year', 'issn', 'doi', 'url', 'keywords',
                 'abstract'],
                {'author', 'title', 'booktitle', 'year', 'doi'}
            ),
            'misc': (
                ['author', 'title', 'booktitle', 'publisher', 'year', 'doi', 'url', 'keywords', 'abstract'],
                {'author', 'title', 'year', 'doi'}
            )
        }

    def _map_csv_content_type(self, csv_content_type: str) -> str:
        """CSV Content Type映射到Bib条目类型"""
        csv_type = csv_content_type.lower().strip()
        if 'article' in csv_type:
            return 'article'
        elif 'chapter' in csv_type:
            return 'incollection'
        elif 'conference' in csv_type:
            return 'inproceedings'
        else:
            return 'misc'

    def _sanitize_text(self, text: str) -> str:
        """清洗文本，适配Zotero格式"""
        if not text or str(text).strip().lower() in ['nan', 'none', '', 'n/a']:
            return 'Not Found'
        clean_text = re.sub(r'\s+', ' ', str(text).strip())
        sensitive_chars = {'{': '\\{', '}': '\\}', '\\': '\\\\', '#': '\\#', '$': '\\$', '&': '\\&', '~': '\\~',
                           '_': '\\_', '^': '\\^', '%': '\\%'}
        for char, escaped in sensitive_chars.items():
            clean_text = clean_text.replace(char, escaped)
        return clean_text

    def _format_authors(self, authors_str: str) -> str:
        """作者格式转换：CSV逗号分隔→Bib的and分隔"""
        if not authors_str or authors_str.strip() == 'Not Found':
            return 'Unknown Author'
        authors = authors_str.split(',')
        authors = [a.strip() for a in authors if a.strip()]
        return ' and '.join(authors)

    def _generate_unique_key(self, row: Dict[str, str], entry_type: str, row_idx: int) -> str:
        """生成唯一Key（作者+年份+标题+行号，避免单CSV内重复）"""
        author = self._sanitize_text(row.get('Authors', 'Unknown'))
        author_key = author.split('and')[0].split()[-1] if 'and' in author else author.split()[-1]
        author_key = re.sub(r'[^a-zA-Z0-9]', '', author_key)[:5]
        year = self._sanitize_text(row.get('Publication Year', '0000'))[:4]
        title = self._sanitize_text(row.get('Item Title', 'NoTitle'))
        title_key = ''.join([w[0].upper() for w in title.split()[:2] if w != 'Not'])
        # 加入行号，确保单CSV内Key唯一
        base_key = f"{author_key}{year}{title_key}{row_idx}"
        return base_key

    def _convert_single_csv_to_bib(self, csv_path: str, output_bib_path: str) -> bool:
        """转换单个CSV到指定Bib文件（新增：文件存在检测）"""
        # ========== 核心修改：检测Bib文件是否已存在 ==========
        if os.path.exists(output_bib_path):
            print(f"⏭️  跳过：{output_bib_path}（文件已存在）")
            self.skip_count += 1
            return True  # 跳过不算失败，标记为成功
        # ====================================================

        try:
            # 读取CSV（多编码兼容）
            encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig', 'latin-1']
            csv_rows = []
            for encoding in encodings:
                try:
                    with open(csv_path, 'r', encoding=encoding, newline='') as f:
                        reader = csv.DictReader(f)
                        csv_rows = list(reader)
                    break
                except Exception:
                    continue
            if not csv_rows:
                print(f"⚠️  {os.path.basename(csv_path)}：无有效数据")
                self.fail_count += 1
                return False

            # 转换每一行
            bib_entries = []
            for row_idx, row in enumerate(csv_rows, 1):
                try:
                    # 确定条目类型
                    csv_content_type = row.get('Content Type', '')
                    entry_type = self._map_csv_content_type(csv_content_type)
                    fields_order, required_fields = self.ref_templates.get(entry_type, self.ref_templates['article'])

                    # 匹配字段
                    bib_fields = []
                    for field in fields_order:
                        # 找到CSV对应的列
                        csv_col = [k for k, v in self.csv_bib_mapping.items() if v == field]
                        csv_value = row.get(csv_col[0], '') if csv_col else ''

                        if field == 'author':
                            value = self._format_authors(self._sanitize_text(csv_value))
                        else:
                            value = self._sanitize_text(csv_value)

                        # 必填字段缺失标注
                        if field in required_fields and value == 'Not Found':
                            bib_fields.append(f"    {field} = {{{value}}}  # 【必填字段缺失】")
                        else:
                            bib_fields.append(f"    {field} = {{{value}}}")

                    # 生成Key并拼接条目
                    entry_key = self._generate_unique_key(row, entry_type, row_idx)
                    fields_str = ',\n'.join(bib_fields)
                    bib_entry = f"@{entry_type}{{{entry_key},\n{fields_str}\n}}"
                    bib_entries.append(bib_entry)
                except Exception as e:
                    print(f"  ⚠️  跳过{os.path.basename(csv_path)}第{row_idx}行：{str(e)}")

            # 写入Bib文件
            if bib_entries:
                os.makedirs(os.path.dirname(output_bib_path), exist_ok=True)
                with open(output_bib_path, 'w', encoding='utf-8') as f:
                    f.write('\n\n'.join(bib_entries))
                self.success_count += 1
                self.generated_files.append(output_bib_path)
                print(f"✅ 生成成功：{output_bib_path}（共{len(bib_entries)}条）")
                return True
            else:
                print(f"⚠️  {os.path.basename(csv_path)}：未生成任何Bib条目")
                self.fail_count += 1
                return False
        except Exception as e:
            print(f"❌ 转换失败：{os.path.basename(csv_path)} - {str(e)}")
            self.fail_count += 1
            return False

    def batch_convert(self, ref_path: str, csv_root_dir: str, output_root_dir: str) -> None:
        """批量转换：遍历所有CSV，生成对应Bib（保留目录结构）"""
        print("📌 单CSV→单Bib转换工具（适配Zotero，带文件存在检测）")
        # 解析参考文件
        self._parse_reference_file(ref_path)

        # 递归扫描所有CSV
        csv_files = []
        for dirpath, _, filenames in os.walk(csv_root_dir):
            for fn in filenames:
                if fn.lower().endswith('.csv'):
                    csv_files.append(os.path.join(dirpath, fn))

        if not csv_files:
            print("⚠️  未找到任何CSV文件")
            return
        print(f"\n🔍 共找到{len(csv_files)}个CSV文件，开始逐个转换...")

        # 逐个转换
        for csv_path in csv_files:
            # 构建输出Bib路径（保留原目录结构）
            relative_path = os.path.relpath(csv_path, csv_root_dir)
            bib_filename = os.path.splitext(relative_path)[0] + '.bib'
            output_bib_path = os.path.join(output_root_dir, bib_filename)

            # 转换单个CSV
            self._convert_single_csv_to_bib(csv_path, output_bib_path)

        # 输出统计结果（新增：跳过数）
        print(f"\n📊 转换完成！")
        print(f"   ✅ 新建成功：{self.success_count} 个文件")
        print(f"   ⏭️  已存在跳过：{self.skip_count} 个文件")
        print(f"   ❌ 转换失败：{self.fail_count} 个文件")
        print(f"\n📁 本次新建的Bib文件列表：")
        for bib_path in self.generated_files:
            print(f"   - {bib_path}")


# 可视化选择函数
def select_reference_file():
    """选择参考Bib/TXT文件"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="第一步：选择Zotero兼容的参考文件（txt/bib）",
        filetypes=[("参考文件", "*.txt *.bib"), ("所有文件", "*.*")]
    )
    root.destroy()
    return path


def select_csv_root_dir():
    """选择CSV根目录"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(
        title="第二步：选择CSV文件所在根目录（自动扫描子目录）",
        mustexist=True
    )
    root.destroy()
    return path


def select_output_root_dir():
    """选择Bib输出根目录（保留原CSV目录结构）"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(
        title="第三步：选择Bib文件输出根目录（会保留原CSV的目录结构）",
        mustexist=True
    )
    root.destroy()
    return path


def main():
    # 1. 选择参考文件
    ref_file = select_reference_file()
    if not ref_file:
        messagebox.showwarning("警告", "未选择参考文件，程序退出")
        return

    # 2. 选择CSV根目录
    csv_dir = select_csv_root_dir()
    if not csv_dir:
        messagebox.showwarning("警告", "未选择CSV目录，程序退出")
        return

    # 3. 选择输出根目录
    output_dir = select_output_root_dir()
    if not output_dir:
        messagebox.showwarning("警告", "未选择输出目录，程序退出")
        return

    # 4. 执行转换
    converter = SingleCsvToBib()
    converter.batch_convert(ref_file, csv_dir, output_dir)

    # 提示完成（新增：跳过数展示）
    messagebox.showinfo(
        "完成",
        f"转换完成！\n✅ 新建成功：{converter.success_count} 个\n⏭️  已存在跳过：{converter.skip_count} 个\n❌ 转换失败：{converter.fail_count} 个"
    )


if __name__ == "__main__":
    main()