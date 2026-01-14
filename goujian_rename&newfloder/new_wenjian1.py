import os

def create_folders_from_lines():
    """
    核心功能：输入多行文本（一行一个文件夹名），按两次回车生成对应文件夹
    适配路径含&等特殊字符，创建后直接列出结果，一目了然
    """
    # 1. 欢迎提示 + 明确目标目录
    target_path = os.path.abspath(os.getcwd())  # 强制使用绝对路径，避免歧义
    print("=" * 60)
    print("📂 多行文件夹快速创建工具")
    print("✅ 用法：粘贴/输入多行名称 → 按两次回车 → 自动创建")
    print(f"📁 文件夹将创建在：\n   {target_path}")
    print("=" * 60)

    # 2. 读取输入（一行一个，空行结束）
    print("\n请输入文件夹名称（一行一个），输完按【两次回车】：")
    folder_names = []
    while True:
        line = input().strip()
        if not line:  # 第二次回车输入空行，停止读取
            break
        folder_names.append(line)

    # 3. 过滤无效输入
    folder_names = list(filter(None, folder_names))  # 去重空值
    if not folder_names:
        print("❌ 未检测到有效文件夹名称！")
        return

    # 4. 展示待创建列表
    print(f"\n📋 待创建文件夹（共{len(folder_names)}个）：")
    for i, name in enumerate(folder_names, 1):
        print(f"  {i}. {name}")

    # 5. 确认创建（避免误操作）
    input("\n👉 按【回车】开始创建（按Ctrl+C取消）...")

    # 6. 核心创建逻辑（兼容特殊路径，无冗余参数）
    print("\n" + "-" * 60)
    print("🚀 正在创建...")
    success = []
    exist = []

    for name in folder_names:
        full_path = os.path.join(target_path, name)
        try:
            if os.path.isdir(full_path):
                exist.append(name)
                continue
            os.makedirs(full_path)  # Windows原生支持，无需额外参数
            success.append(name)
        except Exception as e:
            print(f"❌ {name} 创建失败：{str(e)}")

    # 7. 结果汇总 + 直观验证
    print("\n🎉 创建完成！")
    print(f"✅ 新创建：{len(success)}个 → {', '.join(success)}")
    if exist:
        print(f"ℹ️  已存在：{len(exist)}个 → {', '.join(exist)}")

    # 关键：列出目标目录下的所有文件夹，让你直接看到结果
    print(f"\n🔍 目标目录下的所有文件夹：")
    all_folders = [f for f in os.listdir(target_path) if os.path.isdir(os.path.join(target_path, f))]
    for i, f in enumerate(all_folders, 1):
        print(f"  {i}. {f}")
    print("\n💡 直接复制目标目录路径，粘贴到文件资源管理器打开即可！")
    print("=" * 60)

if __name__ == "__main__":
    create_folders_from_lines()