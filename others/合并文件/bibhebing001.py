import csv
import os
import re
import tkinter as tk
from tkinter import filedialog
from typing import Dict, List, Tuple, Set


class RefDrivenCsvToBib:
    """基于参考Bib结构的CSV转Bib工具（适配Zotero导入）"""

    def __init__(self):
        # 从参考Bib中提取的模板：{条目类型: (字段顺序列表, 必填字段集合)}
        self.ref_bib_templates: Dict[str, Tuple[List[str], Set[str]]] = {}
        # 存储最终合并的Bib条目（去重）
        self.final_bib_entries: Dict[str, str] = {}
        # 扫描到的CSV文件
        self.csv_files: List[str] = []

    def _parse_reference_bib(self, ref_bib_path: str) -> bool:
        """解析参考Bib文件，生成字段模板（核心步骤）"""
        try:
            with open(ref_bib_path, 'r', encoding='utf-8') as f:
                bib_content = f.read()

            # 正则匹配所有Bib条目（支持多行、注释）
            entry_pattern = re.compile(
                r'@(?P<type>\w+)\s*{\s*(?P<key>[^,]+)\s*,[\s\S]*?^\s*}',
                re.MULTILINE | re.IGNORECASE
            )
            field_pattern = re.compile(r'^\s*(?P<field>\w+)\s*=\s*\{.*?\}', re.MULTILINE)

            # 提取每种条目类型的字段顺序和必填字段
            entry_types_seen: Set[str] = set()
            for match in entry_pattern.finditer(bib_content):
                entry_type = match.group('type').lower()
                if entry_type in entry_types_seen:
                    continue  # 同类型只取第一个条目当模板（确保一致性）
                entry_types_seen.add(entry_type)

                # 提取该条目的所有字段（保持原始顺序）
                entry_content = match.group(0)
                field_matches = field_pattern.finditer(entry_content)
                fields_order = [m.group('field').lower() for m in field_matches]

                # 过滤无效字段，保留核心字段（排除key和空字段）
                valid_fields = [f for f in fields_order if f and f not in ['key', 'entrytype']]
                if not valid_fields:
                    continue

                # 标记必填字段（参考Bib中该类型条目必含的字段）
                required_fields = set(valid_fields[:3])  # 前3个字段默认视为必填（参考常规Bib结构）
                self.ref_bib_templates[entry_type] = (valid_fields, required_fields)

            # 验证模板是否提取成功
            if not self.ref_bib_templates:
                print("⚠️  参考Bib解析失败：未找到有效条目，将使用默认模板")
                self._set_default_template()  # 兜底默认模板
                return False
            else:
                print(f"✅ 参考Bib解析成功！提取到{len(self.ref_bib_templates)}种条目模板：")
                for entry_type, (fields, required) in self.ref_bib_templates.items():
                    print(f"  - {entry_type}: 字段顺序={fields}, 必填字段={required}")
                return True
        except Exception as e:
            print(f"❌ 参考Bib读取错误：{str(e)}，将使用默认模板")
            self._set_default_template()
            return False

    def _set_default_template(self):
        """默认模板（当参考Bib解析失败时兜底，适配Zotero常规要求）"""
        self.ref_bib_templates = {
            'article': (['author', 'year', 'title', 'journal', 'volume', 'number', 'pages', 'doi'],
                        {'author', 'year', 'title', 'journal'}),
            'book': (['author', 'year', 'title', 'publisher', 'address', 'doi'],
                     {'author', 'year', 'title', 'publisher'}),
            'inproceedings': (['author', 'year', 'title', 'booktitle', 'publisher', 'pages', 'doi'],
                              {'author', 'year', 'title', 'booktitle'})
        }

    def _sanitize_text(self, text: str) -> str:
        """按参考Bib风格清洗文本（避免Zotero特殊字符报错）"""
        if not text or str(text).strip().lower() in ['nan', 'none', '']:
            return 'Not Found'  # 缺失字段统一标注，避免空字段报错
        # 保留参考Bib的特殊字符处理风格（如大括号、转义符）
        clean_text = re.sub(r'\s+', ' ', str(text).strip())
        # 转义Zotero敏感字符
        for char in ['{', '}', '\\', '#', '$', '&', '~', '_', '^', '%']:
            if char in clean_text:
                clean_text = clean_text.replace(char, f'\\{char}')
        return clean_text

    def _get_entry_template(self, entry_type: str) -> Tuple[List[str], Set[str]]:
        """根据条目类型获取模板（无匹配时用默认article模板）"""
        entry_type = entry_type.lower()
        return self.ref_bib_templates.get(entry_type, self.ref_bib_templates.get('article'))

    def _generate_unique_key(self, row: Dict[str, str], entry_type: str) -> str:
        """生成Zotero兼容的唯一Key（参考Bib风格）"""
        author = self._sanitize_text(row.get('authors', 'Unknown'))
        author_key = author.split(',')[0].split(' ')[-1] if author != 'Not Found' else 'Unknown'
        year = self._sanitize_text(row.get('year', '0000'))[:4]
        title = self._sanitize_text(row.get('title', 'NoTitle'))
        title_key = ''.join([w[0].upper() for w in title.split()[:3] if w != 'Not'])
        # 按参考Bib Key风格拼接（如AuthorYearTitle）
        base_key = f"{author_key}{year}{title_key}"
        final_key = base_key
        counter = 1
        while final_key in self.final_bib_entries:
            final_key = f"{base_key}{counter}"
            counter += 1
        return final_key

    def _csv_to_bib_by_template(self, row: Dict[str, str]) -> str:
        """按参考Bib模板将CSV行转为Bib条目（字段顺序、必填字段严格匹配）"""
        # 1. 确定条目类型（优先CSV的Document Type，无则自动判断）
        doc_type = self._sanitize_text(row.get('document type', ''))
        if 'article' in doc_type.lower() or 'journal' in doc_type.lower():
            entry_type = 'article'
        elif 'book' in doc_type.lower():
            entry_type = 'book'
        elif 'proceeding' in doc_type.lower() or 'conference' in doc_type.lower():
            entry_type = 'inproceedings'
        else:
            entry_type = 'article'  # 默认类型

        # 2. 获取该类型的参考模板（字段顺序+必填字段）
        fields_order, required_fields = self._get_entry_template(entry_type)

        # 3. 按模板顺序匹配CSV字段（CSV列名不区分大小写）
        csv_fields_lower = {k.lower(): v for k, v in row.items()}
        bib_fields = []
        for template_field in fields_order:
            # 匹配CSV中对应的字段（如模板字段author匹配CSV的Authors/author）
            csv_value = csv_fields_lower.get(template_field, csv_fields_lower.get(template_field.replace('_', ''), ''))
            sanitized_value = self._sanitize_text(csv_value)
            # 必填字段缺失时标注警告（便于后续检查）
            if template_field in required_fields and sanitized_value == 'Not Found':
                bib_fields.append(f"    {template_field} = {{{sanitized_value}}}  # 必填字段缺失！")
            else:
                bib_fields.append(f"    {template_field} = {{{sanitized_value}}}")

        # 4. 生成唯一Key并拼接条目
        entry_key = self._generate_unique_key(row, entry_type)
        fields_str = ',\n'.join(bib_fields)
        return f"@{entry_type}{{{entry_key},\n{fields_str}\n}}"

    def _scan_csv_dir(self, root_dir: str) -> None:
        """递归扫描CSV目录及子目录"""
        self.csv_files = []
        for dirpath, _, filenames in os.walk(root_dir):
            for fn in filenames:
                if fn.lower().endswith('.csv'):
                    self.csv_files.append(os.path.join(dirpath, fn))
        print(f"\n🔍 扫描到{len(self.csv_files)}个CSV文件（含子目录）")

    def _convert_single_csv(self, csv_path: str) -> int:
        """按参考模板转换单个CSV，返回成功条数"""
        success = 0
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1', 'utf-8-sig']  # 多编码兼容
        for encoding in encodings:
            try:
                with open(csv_path, 'r', encoding=encoding, newline='') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        try:
                            bib_entry = self._csv_to_bib_by_template(row)
                            entry_key = bib_entry.split('{')[1].split(',')[0]  # 提取Key
                            self.final_bib_entries[entry_key] = bib_entry
                            success += 1
                        except Exception as e:
                            print(f"  ⚠️  跳过CSV行：{str(e)}")
                print(f"✅ 转换完成：{os.path.basename(csv_path)}（成功{success}条，编码{encoding}）")
                return success
            except Exception as e:
                continue
        print(f"❌ 读取失败：{os.path.basename(csv_path)}（所有编码尝试无效）")
        return 0

    def batch_convert_and_merge(self, ref_bib_path: str, csv_root_dir: str, output_bib_path: str) -> None:
        """主流程：解析参考Bib→扫描CSV→批量转换→合并输出"""
        print("📌 参考Bib驱动的CSV转Bib工具（适配Zotero）")
        # 1. 解析参考Bib
        self._parse_reference_bib(ref_bib_path)
        # 2. 扫描CSV目录
        self._scan_csv_dir(csv_root_dir)
        if not self.csv_files:
            print("⚠️  未找到任何CSV文件，程序退出")
            return
        # 3. 批量转换CSV
        total_success = 0
        for csv_path in self.csv_files:
            total_success += self._convert_single_csv(csv_path)
        # 4. 合并并写入最终Bib（按Key排序，Zotero导入更友好）
        if self.final_bib_entries:
            sorted_entries = sorted(self.final_bib_entries.values(), key=lambda x: x.split('{')[1].split(',')[0])
            with open(output_bib_path, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(sorted_entries))
            print(f"\n🎉 全部完成！")
            print(
                f"📊 统计：{len(self.csv_files)}个CSV → 成功转换{total_success}条 → 合并后{len(self.final_bib_entries)}个唯一条目")
            print(f"📁 输出文件：{output_bib_path}")
        else:
            print("⚠️  未生成任何Bib条目")


def select_reference_bib():
    """选择正确的参考Bib文件"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="第一步：选择Zotero可正常导入的参考Bib文件",
        filetypes=[("BibTeX文件", "*.bib"), ("所有文件", "*.*")]
    )
    root.destroy()
    return path


def select_csv_root_dir():
    """选择CSV根目录（含子目录）"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(
        title="第二步：选择CSV文件所在的根目录（自动扫描子目录）",
        mustexist=True
    )
    root.destroy()
    return path


def select_output_bib():
    """选择最终输出的Bib文件路径"""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.asksaveasfilename(
        title="第三步：选择合并后Bib文件的保存位置",
        defaultextension=".bib",
        filetypes=[("BibTeX文件", "*.bib"), ("所有文件", "*.*")],
        initialfile="zotero_compatible_bib.bib"
    )
    root.destroy()
    return path


def main():
    # 1. 选择参考Bib
    ref_bib = select_reference_bib()
    if not ref_bib:
        print("⚠️  未选择参考Bib文件，程序退出")
        return
    # 2. 选择CSV根目录
    csv_dir = select_csv_root_dir()
    if not csv_dir:
        print("⚠️  未选择CSV目录，程序退出")
        return
    # 3. 选择输出Bib
    output_bib = select_output_bib()
    if not output_bib:
        print("⚠️  未选择输出文件，程序退出")
        return
    # 4. 执行转换+合并
    converter = RefDrivenCsvToBib()
    converter.batch_convert_and_merge(ref_bib, csv_dir, output_bib)


if __name__ == "__main__":
    main()