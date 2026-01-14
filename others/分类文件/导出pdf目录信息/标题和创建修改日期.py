import os
import csv
import platform
from datetime import datetime


def get_file_info(file_path):
    """
    获取文件的详细信息（大小、创建/修改时间）
    :param file_path: 文件的绝对路径
    :return: 包含文件信息的字典
    """
    try:
        # 获取文件基本属性
        file_stat = os.stat(file_path)

        # 兼容不同系统的创建时间
        if platform.system() == 'Windows':
            create_time = datetime.fromtimestamp(file_stat.st_ctime)
        else:
            # macOS/Linux 优先用birthtime，无则用ctime
            create_time = datetime.fromtimestamp(
                file_stat.st_birthtime if hasattr(file_stat, 'st_birthtime') else file_stat.st_ctime)

        # 转换文件大小（字节→MB，保留2位小数）
        file_size_mb = round(file_stat.st_size / (1024 * 1024), 2)

        return {
            '文件名': os.path.basename(file_path),
            '文件路径': os.path.abspath(file_path),
            '文件大小(MB)': file_size_mb,
            '创建时间': create_time.strftime('%Y-%m-%d %H:%M:%S'),
            '修改时间': datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        }
    except Exception as e:
        return {
            '文件名': os.path.basename(file_path),
            '文件路径': os.path.abspath(file_path),
            '文件大小(MB)': '获取失败',
            '创建时间': '获取失败',
            '修改时间': '获取失败',
            '错误信息': str(e)
        }


def scan_pdf_files(folder_path):
    """
    扫描指定文件夹下的所有PDF文件，提取完整信息
    :param folder_path: 文件夹路径
    :return: PDF文件信息列表
    """
    pdf_files_info = []

    # 检查文件夹是否存在
    if not os.path.isdir(folder_path):
        print(f"❌ 错误：文件夹 '{folder_path}' 不存在！")
        return pdf_files_info

    # 筛选PDF文件
    pdf_files = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path) and file_name.lower().endswith('.pdf'):
            pdf_files.append(file_path)

    if not pdf_files:
        print("⚠️  该文件夹下未找到任何PDF文件！")
        return pdf_files_info

    # 提取每个PDF的详细信息
    total_files = len(pdf_files)
    print(f"\n✅ 找到 {total_files} 个PDF文件，开始提取信息...")
    for idx, file_path in enumerate(pdf_files, 1):
        print(f"正在处理 [{idx}/{total_files}]：{os.path.basename(file_path)}")
        file_info = get_file_info(file_path)
        pdf_files_info.append(file_info)

    return pdf_files_info


def export_to_csv(file_info_list, output_path='pdf_files_full_info.csv'):
    """
    将PDF完整信息导出为CSV文件
    """
    if not file_info_list:
        return

    # 获取所有字段名（兼容有/无错误信息的情况）
    fieldnames = []
    for info in file_info_list:
        fieldnames.extend(info.keys())
    fieldnames = list(set(fieldnames))  # 去重
    # 固定字段顺序，提升可读性
    fixed_order = ['文件名', '文件路径', '文件大小(MB)', '创建时间', '修改时间', '错误信息']
    fieldnames = [f for f in fixed_order if f in fieldnames]

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(file_info_list)

    print(f"\n📄 完整信息已导出到：{os.path.abspath(output_path)}")


def main():
    print("=" * 60)
    print("        PDF文件完整信息提取工具")
    print("=" * 60)

    # 交互式输入文件夹路径
    while True:
        folder_path = input("\n请输入PDF文件所在的文件夹路径（例如：D:/我的文档/PDF文件）：").strip()

        if not folder_path:
            print("⚠️  路径不能为空，请重新输入！")
            continue

        folder_path = os.path.normpath(folder_path)

        if os.path.isdir(folder_path):
            break
        else:
            print(f"❌ 路径 '{folder_path}' 不存在或不是文件夹，请重新输入！")

    # 扫描并提取PDF信息
    print(f"\n🔍 正在扫描文件夹：{os.path.abspath(folder_path)}")
    pdf_info = scan_pdf_files(folder_path)

    # 导出CSV
    export_to_csv(pdf_info)

    # 完成提示
    print("\n🎉 操作完成！按任意键退出...")
    os.system("pause >nul 2>&1")


if __name__ == '__main__':
    main()