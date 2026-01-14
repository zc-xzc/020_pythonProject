import os
import csv
import pdfplumber


def extract_pdf_text(pdf_path):
    """
    提取PDF文件的文本内容
    :param pdf_path: PDF文件路径
    :return: 提取的文本内容（失败返回错误信息）
    """
    try:
        text_content = ""
        # 打开PDF文件并提取所有页面文本
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # 提取单页文本，过滤空行
                page_text = page.extract_text() or ""
                text_content += page_text.strip() + "\n\n"

        # 去除多余空白，保留主要内容
        text_content = '\n'.join([line.strip() for line in text_content.split('\n') if line.strip()])
        return text_content if text_content else "PDF无文本内容（可能是图片型PDF）"

    except Exception as e:
        return f"提取失败：{str(e)}（可能是加密PDF/损坏PDF/图片型PDF）"


def scan_pdf_with_content(folder_path):
    """
    扫描指定文件夹下的PDF文件，提取文件名和文本内容
    :param folder_path: 文件夹路径
    :return: 包含文件名和内容的列表
    """
    pdf_info_list = []

    # 检查文件夹是否存在
    if not os.path.isdir(folder_path):
        print(f"❌ 错误：文件夹 '{folder_path}' 不存在！")
        return pdf_info_list

    # 筛选文件夹中的PDF文件
    pdf_files = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path) and file_name.lower().endswith('.pdf'):
            pdf_files.append((file_name, file_path))

    if not pdf_files:
        print("⚠️  该文件夹下未找到任何PDF文件！")
        return pdf_info_list

    # 遍历提取每个PDF的内容
    total_files = len(pdf_files)
    print(f"\n✅ 找到 {total_files} 个PDF文件，开始提取内容...")
    for idx, (file_name, file_path) in enumerate(pdf_files, 1):
        print(f"\n正在处理 [{idx}/{total_files}]：{file_name}")

        # 提取文本内容
        text_content = extract_pdf_text(file_path)

        # 保存信息
        pdf_info_list.append({
            '文件名': file_name,
            '文件路径': os.path.abspath(file_path),
            '文本内容': text_content[:10000]  # 限制内容长度，避免CSV文件过大
        })

    return pdf_info_list


def export_to_csv(pdf_info_list, output_path='pdf_names_content.csv'):
    """
    导出文件名和内容到CSV文件
    """
    if not pdf_info_list:
        return

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
        # utf-8-sig 确保Excel打开中文不乱码
        writer = csv.DictWriter(csvfile, fieldnames=['文件名', '文件路径', '文本内容'])
        writer.writeheader()
        writer.writerows(pdf_info_list)

    print(f"\n📄 结果已导出到：{os.path.abspath(output_path)}")


def main():
    print("=" * 60)
    print("        PDF文件名+内容提取工具")
    print("=" * 60)

    # 交互式输入文件夹路径（支持多次输入，直到路径合法）
    while True:
        folder_path = input("\n请输入PDF文件所在的文件夹路径（例如：D:/我的文档/PDF文件）：").strip()

        # 处理空输入
        if not folder_path:
            print("⚠️  路径不能为空，请重新输入！")
            continue

        # 格式化路径（处理用户输入的反斜杠/斜杠混用问题）
        folder_path = os.path.normpath(folder_path)

        # 检查路径是否存在
        if os.path.isdir(folder_path):
            break
        else:
            print(f"❌ 路径 '{folder_path}' 不存在或不是文件夹，请重新输入！")

    # 扫描并提取PDF信息
    pdf_info = scan_pdf_with_content(folder_path)

    # 导出CSV
    export_to_csv(pdf_info)

    print("\n🎉 操作完成！按任意键退出...")
    # 防止命令行窗口直接关闭（Windows）
    os.system("pause >nul 2>&1")


if __name__ == '__main__':
    main()