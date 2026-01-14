import csv
import os
import re
import tkinter as tk
from tkinter import filedialog
from typing import Dict, List, Tuple, Set


class ZoteroCsvToBib:
    """适配Zotero的CSV转Bib工具（基于参考Bib格式）"""

    def __init__(self):
        # 参考Bib字段模板（从你的参考文件中提取，优先级最高）
        self.ref_templates: Dict[str, Tuple[List[str], Set[str]]] = {}
        # CSV列名→Bib字段映射（适配你的SearchResults.csv）
        self.csv_bib_mapping = {
            'Item Title': 'title',
            'Authors': 'author',
            'Publication Year': 'year',
            'Publication Title': 'journal',  # Article的期刊名
            'Book Series Title': 'booktitle',  # Chapter/会议集的书名
            'Journal Volume': 'volume',
            'Journal Issue': 'number',
            'Pages': 'pages',
            'Item DOI': 'doi',
            'URL': 'url',
            'Content Type': 'entrytype',
            'Editor': 'editor'  # CSV中若有编辑列可补充，无则自动提取
        }
        # 存储合并后的Bib条目（去重）
        self.final_entries: Dict[str, str] = {}
        # 扫描到的CSV文件
        self.csv_files: List[str] = []

    def _parse_reference_file(self, ref_path: str) -> bool:
        """解析参考文件（忽略后缀，只看内容），提取字段模板"""
        try:
            # 支持txt/bib后缀，读取内容
            with open(ref_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 正则匹配所有Bib条目（支持多行、空格变化）
            entry_pattern = re.compile(
                r'@(?P<type>\w+)\s*{\s*(?P<key>[^,]+)\s*,[\s\S]*?^\s*}',
                re.MULTILINE | re.IGNORECASE
            )
            field_pattern = re.compile(r'^\s*(?P<field>\w+)\s*=\s*\{(?P<value>.*?)\}', re.MULTILINE | re.DOTALL)

            for match in entry_pattern.finditer(content):
                entry_type = match.group('type').lower()
                if entry_type in self.ref_templates:
                    continue  # 同类型只保留第一个模板（保证一致性）

                # 提取该条目的字段顺序和必填字段（前5个字段视为必填）
                entry_content = match.group(0)
                field_matches = field_pattern.finditer(entry_content)
                fields_order = [m.group('field').lower() for m in field_matches if
                                m.group('field').lower() not in ['key']]
                required_fields = set(fields_order[:5])  # 参考Bib前5个字段为必填
                self.ref_templates[entry_type] = (fields_order, required_fields)

            # 验证模板提取结果
            if not self.ref_templates:
                print("⚠️  参考文件解析失败，使用默认Zotero兼容模板")
                self._set_default_template()
                return False
            else:
                print(f"✅ 参考文件解析成功！提取{len(self.ref_templates)}种条目模板：")
                for t, (fields, req) in self.ref_templates.items():
                    print(f"  - {t}: 字段顺序={fields[:8]}...  必填字段={req}")
                return True
        except Exception as e:
            print(f"❌ 参考文件读取错误：{str(e)}，使用默认模板")
            self._set_default_template()
            return False

    def _set_default_template(self):
        """默认模板（Zotero兼容版，基于参考Bib风格）"""
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
        """CSV的Content Type映射到Bib条目类型（适配你的表格）"""
        csv_type = csv_content_type.lower().strip()
        if 'article' in csv_type:
            return 'article'
        elif 'chapter' in csv_type:
            return 'incollection'
        elif 'conference' in csv_type:
            return 'inproceedings'
        else:
            return 'misc'  # 其他类型默认misc

    def _sanitize_text(self, text: str) -> str:
        """按参考Bib风格清洗文本（保留特殊字符，处理空值）"""
        if not text or str(text).strip().lower() in ['nan', 'none', '', 'n/a']:
            return 'Not Found'
        # 保留原始特殊字符（如μM、°C），只转义Bib敏感字符
        clean_text = re.sub(r'\s+', ' ', str(text).strip())
        sensitive_chars = {'{': '\\{', '}': '\\}', '\\': '\\\\', '#': '\\#', '$': '\\$', '&': '\\&', '~': '\\~',
                           '_': '\\_', '^': '\\^', '%': '\\%'}
        for char, escaped in sensitive_chars.items():
            clean_text = clean_text.replace(char, escaped)
        return clean_text

    def _format_authors(self, authors_str: str) -> str:
        """格式化作者：CSV中的逗号分隔→Bib的and分隔（参考Bib风格）"""
        if not authors_str or authors_str.strip() == 'Not Found':
            return 'Unknown Author'
        # CSV中作者格式如"Vir SinghR. K. Srivastava"→拆分并以and连接
        # 处理可能的姓名连写（按大写字母拆分）
        authors = re.split(r'(?=[A-Z])', authors_str.strip())
        authors = [a.strip() for a in authors if a.strip()]
        # 若拆分失败，直接按逗号分隔
        if len(authors) < 2:
            authors = authors_str.split(',')
            authors = [a.strip() for a in authors if a.strip()]
        return ' and '.join(authors)

    def _generate_unique_key(self, row: Dict[str, str], entry_type: str) -> str:
        """生成Zotero兼容Key（参考Bib风格：作者首字母+年份+标题首字母）"""
        author = self._sanitize_text(row.get('Authors', 'Unknown'))
        author_key = author.split('and')[0].split()[-1] if 'and' in author else author.split()[-1]
        author_key = re.sub(r'[^a-zA-Z0-9]', '', author_key)[:5]  # 取前5个字母/数字
        year = self._sanitize_text(row.get('Publication Year', '0000'))[:4]
        title = self._sanitize_text(row.get('Item Title', 'NoTitle'))
        title_key = ''.join([w[0].upper() for w in title.split()[:3] if w != 'Not'])
        base_key = f"{author_key}{year}{title_key}"
        final_key = base_key
        counter = 1
        while final_key in self.final_entries:
            final_key = f"{base_key}{counter}"
            counter += 1
        return final_key

    def _csv_row_to_bib(self, row: Dict[str, str]) -> str:
        """按参考Bib模板转换单行CSV→Bib条目（字段顺序、格式严格匹配）"""
        # 1. 确定条目类型
        csv_content_type = row.get('Content Type', '')
        entry_type = self._map_csv_content_type(csv_content_type)
        # 获取该类型的参考模板（无则用article模板）
        fields_order, required_fields = self.ref_templates.get(entry_type, self.ref_templates['article'])

        # 2. 匹配CSV数据到Bib字段
        bib_data: Dict[str, str] = {}
        for csv_col, bib_field in self.csv_bib_mapping.items():
            if bib_field == 'entrytype':
                continue
            csv_value = row.get(csv_col, '')
            if bib_field == 'author':
                bib_data[bib_field] = self._format_authors(self._sanitize_text(csv_value))
            else:
                bib_data[bib_field] = self._sanitize_text(csv_value)

        # 3. 按参考模板顺序构建字段行（缺失字段标注，必填字段高亮警告）
        bib_fields = []
        for field in fields_order:
            value = bib_data.get(field, 'Not Found')
            # 必填字段缺失标注警告
            if field in required_fields and value == 'Not Found':
                bib_fields.append(f"    {field} = {{{value}}}  # 【必填字段缺失】")
            else:
                bib_fields.append(f"    {field} = {{{value}}}")

        # 4. 生成Key并拼接条目
        entry_key = self._generate_unique_key(row, entry_type)
        fields_str = ',\n'.join(bib_fields)
        return f"@{entry_type}{{{entry_key},\n{fields_str}\n}}"

    def _scan_csv_dir(self, root_dir: str) -> None:
        """递归扫描目录及子目录下的所有CSV文件"""
        self.csv_files = []
        for dirpath, _, filenames in os.walk(root_dir):
            for fn in filenames:
                if fn.lower().endswith('.csv'):
                    self.csv_files.append(os.path.join(dirpath, fn))
        print(f"\n🔍 扫描完成：找到{len(self.csv_files)}个CSV文件（含子目录）")

    def _convert_single_csv(self, csv_path: str) -> int:
        """转换单个CSV，返回成功条数"""
        success = 0
        # 多编码兼容（UTF-8/GBK/GB2312）
        encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig', 'latin-1']
        for encoding in encodings:
            try:
                with open(csv_path, 'r', encoding=encoding, newline='') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        try:
                            bib_entry = self._csv_row_to_bib(row)
                            entry_key = bib_entry.split('{')[1].split(',')[0]
                            self.final_entries[entry_key] = bib_entry
                            success += 1
                        except Exception as e:
                            print(f"  ⚠️  跳过CSV行：{str(e)}")
                print(f"✅ 转换成功：{os.path.basename(csv_path)}（编码：{encoding}，成功{success}条）")
                return success
            except Exception as e:
                continue
        print(f"❌ 读取失败：{os.path.basename(csv_path)}（所有编码尝试无效）")
        return 0

    def batch_convert_and_merge(self, ref_path: str, csv_root_dir: str, output_path: str) -> None:
        """主流程：解析参考文件→扫描CSV→转换→合并输出"""
        print("📌 Zotero兼容版CSV转Bib工具（严格匹配参考格式）")
        # 1. 解析参考文件
        self._parse_reference_file(ref_path)
        # 2. 扫描CSV目录
        self._scan_csv_dir(csv_root_dir)
        if not self.csv_files:
            print("⚠️  未找到CSV文件，程序退出")
            return
        # 3. 批量转换
        total_success = 0
        for csv_file in self.csv_files:
            total_success += self._convert_single_csv(csv_file)
        # 4. 合并输出（按Key排序，Zotero导入更友好）
        if self.final_entries:
            sorted_entries = sorted(self.final_entries.values(), key=lambda x: x.split('{')[1].split(',')[0])
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(sorted_entries))
            print(f"\n🎉 全部完成！")
            print(
                f"📊 统计：{len(self.csv_files)}个CSV → 成功转换{total_success}条 → 合并后{len(self.final_entries)}个唯一条目")
            print(f"📁 输出文件：{output_path}")
        else:
            print("⚠️  未生成任何Bib条目")


# 可视化选择函数
def select_reference_file():
    """选择参考文件（支持txt/bib后缀）"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="第一步：选择Zotero兼容的参考文件（txt/bib格式）",
        filetypes=[("参考文件", "*.txt *.bib"), ("所有文件", "*.*")]
    )
    root.destroy()
    return path


def select_csv_dir():
    """选择CSV根目录（递归扫描子目录）"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(
        title="第二步：选择CSV文件所在根目录（自动扫描子目录）",
        mustexist=True
    )
    root.destroy()
    return path


def select_output_file():
    """选择输出Bib文件路径"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.asksaveasfilename(
        title="第三步：选择输出Bib文件位置",
        defaultextension=".bib",
        filetypes=[("BibTeX文件", "*.bib"), ("所有文件", "*.*")],
        initialfile="zotero_compatible_merged.bib"
    )
    root.destroy()
    return path


def main():
    # 1. 选择参考文件
    ref_file = select_reference_file()
    if not ref_file:
        print("⚠️  未选择参考文件，程序退出")
        return
    # 2. 选择CSV目录
    csv_dir = select_csv_dir()
    if not csv_dir:
        print("⚠️  未选择CSV目录，程序退出")
        return
    # 3. 选择输出文件
    output_file = select_output_file()
    if not output_file:
        print("⚠️  未选择输出文件，程序退出")
        return
    # 4. 执行转换合并
    converter = ZoteroCsvToBib()
    converter.batch_convert_and_merge(ref_file, csv_dir, output_file)


if __name__ == "__main__":
    main()