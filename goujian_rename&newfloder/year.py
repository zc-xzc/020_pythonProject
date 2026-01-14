import os


def create_year_folders(start_year=1990, end_year=2025):
    """
    创建从起始年份到结束年份的所有年份文件夹

    参数:
        start_year: 起始年份（默认1990）
        end_year: 结束年份（默认2025）
    """
    # 获取当前脚本所在目录（也可自定义路径，如：target_path = "D:/年份文件夹"）
    target_path = os.getcwd()

    print(f"开始创建文件夹，目标路径：{target_path}")
    print(f"年份范围：{start_year} - {end_year}")
    print("-" * 50)

    created_count = 0
    existed_count = 0

    # 遍历每个年份，创建对应的文件夹
    for year in range(start_year, end_year + 1):
        # 文件夹名称（纯数字年份，如"1990"）
        folder_name = str(year)
        # 完整路径
        folder_path = os.path.join(target_path, folder_name)

        try:
            # 创建文件夹，如果已存在则不报错（exist_ok=True）
            os.makedirs(folder_path, exist_ok=True)

            if os.path.exists(folder_path):
                print(f"✅ 成功创建/已存在：{folder_name}")
                if os.path.isdir(folder_path):
                    created_count += 1
                else:
                    existed_count += 1
        except Exception as e:
            print(f"❌ 创建失败 {folder_name}：{str(e)}")

    print("-" * 50)
    print(f"任务完成！")
    print(f"总计需要创建：{end_year - start_year + 1} 个文件夹")
    print(f"成功创建/已存在：{created_count} 个")
    print(f"已存在无需创建：{existed_count} 个")


if __name__ == "__main__":
    # 调用函数，创建1990-2025年文件夹
    create_year_folders(1990, 2025)