import os
import shutil
import string
from pdfplumber import open as open_pdf
from typing import List, Tuple

# ===================== 核心关键词配置（中文+英文，包含短语和单词）=====================
# 格式：(中文关键词/短语, 英文关键词/短语，多个用逗号分隔)
KEYWORDS: List[Tuple[str, str]] = [
    ("环境法", "Environmental Law"),
    ("空气法", "Air Law,Clean Air Act"),
    ("土壤法", "Soil Law"),
    ("水法", "Water Law,Clean Water Act"),
    ("能源法", "Energy Law"),
    ("环保法", "Environmental Protection Law"),
    ("环境规制", "Environmental regulation"),
    ("污染防治", "Pollution prevention and control"),
    ("节能", "Energy conservation"),
    ("减排", "Emission reduction"),
    ("排污", "Pollutant discharge"),
    ("污染税", "Pollution tax"),
    ("环境税", "Environmental tax"),
    ("能源税", "Energy tax"),
    ("排污交易", "Pollutant trading,Emissions trading"),
    ("碳税", "Carbon tax"),
    ("碳交易", "Carbon trading,Carbon market"),
    ("废物管理", "Waste management,Solid waste management"),
    ("资源利用", "Resource utilization,Resource use efficiency"),
    ("环境信息披露", "Environmental information disclosure"),
    ("生态补偿", "Ecological compensation,Carbon offset"),
    ("绿色金融", "Green finance"),
]


def preprocess_text(text: str) -> str:
    """文本预处理：统一小写、去除标点、合并多余空格，提升匹配全面性"""
    # 1. 转为小写
    text = text.lower()
    # 2. 去除所有标点符号
    translator = str.maketrans("", "", string.punctuation)
    text = text.translate(translator)
    # 3. 合并多个空格为单个空格，去除首尾空格
    text = " ".join(text.split())
    # 4. 去除中文全角标点（补充处理）
    chinese_punctuation = "，。！？；：""''（）【】《》、·~@#￥%……&*——+=-_()"
    for p in chinese_punctuation:
        text = text.replace(p, "")
    return text


def extract_pdf_text(pdf_path: str) -> str:
    """提取PDF文本，兼容更多格式，确保文本提取全面"""
    text = ""
    try:
        with open_pdf(pdf_path) as pdf:
            # 启用布局分析，提升文本提取准确性（针对有格式的PDF）
            for page in pdf.pages:
                page_text = page.extract_text(layout=True) or ""
                text += page_text
        # 预处理文本，统一格式
        text = preprocess_text(text)
        return text
    except Exception as e:
        print(f"⚠️  读取PDF失败：{pdf_path}，错误信息：{str(e)}")
        return ""


def is_contains_any_keyword(text: str) -> bool:
    """全面匹配关键词：包含短语/单词的任意形式都判定为相关"""
    # 1. 扁平化所有关键词（拆分短语和单词，去重）
    all_keywords_flat = []
    for cn, en in KEYWORDS:
        # 处理中文关键词（短语/单词）
        all_keywords_flat.append(preprocess_text(cn))
        # 处理英文关键词（拆分多个翻译，逐个加入）
        en_keywords = [preprocess_text(e) for e in en.split(",")]
        all_keywords_flat.extend(en_keywords)

    # 去重，避免重复匹配
    all_keywords_flat = list(set(all_keywords_flat))

    # 2. 逐个匹配，只要包含任意一个关键词就返回True
    for keyword in all_keywords_flat:
        if keyword in text:
            print(f"🔍 匹配到关键词：{keyword}")  # 打印匹配到的关键词，方便核对
            return True
    return False


def main():
    # 1. 接收用户输入的文件夹路径（支持拖拽路径）
    print("📌 提示：文件夹路径可直接拖拽到命令行窗口输入")
    input_dir = input("请输入待筛选PDF的文件夹路径：").strip().strip('"')  # 去除路径首尾的引号（拖拽时可能带）
    output_dir = input("请输入保留文件的输出文件夹路径：").strip().strip('"')

    # 2. 验证输入文件夹
    if not os.path.isdir(input_dir):
        print(f"❌ 错误：输入文件夹 {input_dir} 不存在！")
        return

    # 3. 创建输出文件夹（不存在则自动创建）
    os.makedirs(output_dir, exist_ok=True)
    print(f"✅ 输出文件夹已就绪：{output_dir}")

    # 4. 筛选所有PDF文件（忽略大小写）
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print(f"ℹ️  输入文件夹中未找到PDF文件！")
        return

    print(f"\nℹ️  共检测到 {len(pdf_files)} 个PDF文件，开始全面检索关键词...")
    print("-" * 60)

    # 统计变量
    retained = 0  # 保留数
    deleted = 0  # 删除数
    skipped = 0  # 跳过数（加密/损坏）

    for filename in pdf_files:
        file_path = os.path.join(input_dir, filename)
        print(f"\n正在处理：{filename}")

        # 提取并预处理PDF文本
        pdf_text = extract_pdf_text(file_path)
        if not pdf_text:
            print(f"⚠️  文件无法读取（加密/损坏/无文本），跳过处理")
            skipped += 1
            continue

        # 全面匹配关键词
        if is_contains_any_keyword(pdf_text):
            # 保留：复制到输出文件夹（避免重名）
            output_path = os.path.join(output_dir, filename)
            counter = 1
            while os.path.exists(output_path):
                name, ext = os.path.splitext(filename)
                output_path = os.path.join(output_dir, f"{name}_{counter}{ext}")
                counter += 1
            shutil.copy2(file_path, output_path)  # 保留文件元信息
            retained += 1
            print(f"✅ 包含相关关键词，已保留至输出文件夹")
        else:
            # 删除：需手动确认，避免误删
            confirm = input(f"❌ 未匹配到任何关键词，是否删除？(y/n，默认n)：").strip().lower()
            if confirm == "y":
                os.remove(file_path)
                deleted += 1
                print(f"✅ 已删除该文件")
            else:
                print(f"ℹ️  已保留该文件（用户取消删除）")

    # 结果汇总
    print("\n" + "=" * 60)
    print(f"📊 检索完成！汇总结果：")
    print(f"• 总处理文件数：{len(pdf_files)}")
    print(f"• 保留文件数（含关键词）：{retained}")
    print(f"• 删除文件数（无关键词）：{deleted}")
    print(f"• 跳过文件数（无法读取）：{skipped}")
    print(f"• 保留文件路径：{output_dir}")
    print("=" * 60)


if __name__ == "__main__":
    main()