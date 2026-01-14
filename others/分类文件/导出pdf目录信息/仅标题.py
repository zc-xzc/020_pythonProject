import os
import csv


def scan_pdf_filenames(folder_path):
    """
    扫描指定文件夹下的所有PDF文件，仅提取文件名
    :param folder_path: 文件夹路径
    :return: PDF文件名列表
    """
    pdf_filenames = []

    # 检查文件夹是否存在
    if not os.path.isdir(folder_path):
        print(f"❌ 错误：文件夹 '{folder_path}' 不存在！")
        return pdf_filenames

    # 筛选PDF文件并提取名称
    pdf_files = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path) and file_name.lower().endswith('.pdf'):
            pdf_files.append(file_name)

    if not pdf_files:
        print("⚠️  该文件夹下未找到任何PDF文件！")
        return pdf_filenames

    # 整理成字典格式，方便导出CSV
    for name in pdf_files:
        pdf_filenames.append({'文件名': name})

    return pdf_filenames


def export_filenames_to_csv(filenames_list, output_path='pdf_filenames.csv'):
    """
    将PDF文件名导出为CSV文件
    """
    if not filenames_list:
        return

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=['文件名'])
        writer.writeheader()
        writer.writerows(filenames_list)

    print(f"\n📄 文件名列表已导出到：{os.path.abspath(output_path)}")


def main():
    print("=" * 60)
    print("        PDF文件名提取工具")
    print("=" * 60)

    # 交互式输入文件夹路径（循环校验，直到路径合法）
    while True:
        folder_path = input("\n请输入PDF文件所在的文件夹路径（例如：D:/我的文档/PDF文件）：").strip()

        # 处理空输入
        if not folder_path:
            print("⚠️  路径不能为空，请重新输入！")
            continue

        # 格式化路径（兼容/和\混用）
        folder_path = os.path.normpath(folder_path)

        # 检查路径合法性
        if os.path.isdir(folder_path):
            break
        else:
            print(f"❌ 路径 '{folder_path}' 不存在或不是文件夹，请重新输入！")

    # 扫描PDF文件名
    print(f"\n🔍 正在扫描文件夹：{os.path.abspath(folder_path)}")
    pdf_filenames = scan_pdf_filenames(folder_path)

    # 导出CSV
    export_filenames_to_csv(pdf_filenames)

    # 完成提示，防止窗口关闭
    print("\n🎉 操作完成！按任意键退出...")
    os.system("pause >nul 2>&1")


if __name__ == '__main__':
    main()