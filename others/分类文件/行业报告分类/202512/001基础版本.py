import os
import shutil
import time
from collections import defaultdict

# ===================== 【核心配置区】- 可自由扩展/修改 =====================
# 1. 领域-关键词映射（基于你的所有文件名优化，后续新增领域/关键词只改这里）
DOMAIN_KEYWORDS = {
    # 人工智能领域（全覆盖你的AI相关文件名）
    "人工智能": [
        "AI", "人工智能", "大模型", "AIGC", "Agent", "智能体", "算力", "生成式",
        "XR", "脑机接口", "量子计算", "具身智能", "数字员工", "智能算力", "AI办公",
        "AI硬件", "AI学习", "生成式AI", "AI Agent", "智能媒体", "人机友好",
        "中文大模型", "AI应用", "AI+", "人工智能教育", "AI治理", "AI软实力"
    ],
    # 新能源领域（全覆盖你的新能源/电池/光伏相关文件名）
    "新能源": [
        "新能源", "光伏", "锂电", "电池", "碳中和", "储能", "氢能", "碳达峰",
        "液流电池", "绿电", "固态电池", "新能源汽车", "光伏建设", "绿色低碳",
        "能源转型", "电力设备", "电动车", "绿电交易", "能源大数据"
    ],
    # 金融领域（覆盖你的投资/并购/金融科技相关文件名）
    "金融": [
        "金融", "证券", "投资", "并购", "支付", "财富", "保险", "市值", "融资",
        "ESG", "跨境金融", "另类数据", "金融科技", "国资并购", "加密行业",
        "黄金需求", "家庭财富", "央国企金融", "AI金融", "出海金融"
    ],
    # 医疗健康领域（覆盖你的医药/精准医疗/健康消费相关文件名）
    "医疗健康": [
        "医疗", "医药", "健康", "精准医疗", "医健消费", "疫苗", "基因检测",
        "医疗器械", "医药生物", "中药", "轻医美", "心理健康", "养老服务",
        "银发经济", "亚健康", "小核酸", "中药代煎", "医疗投融资"
    ],
    # 消费零售领域（覆盖你的消费/电商/家居/文旅相关文件名）
    "消费零售": [
        "消费", "零售", "电商", "品牌", "跨境电商", "家居", "旅游", "民宿",
        "户外", "多巴胺经济", "Z世代", "05后", "银发消费", "玄学消费",
        "适老化", "零售数字化", "自有品牌", "双十一", "冬季消费", "运动商业",
        "内容社区", "年轻人生活方式", "消费新图景", "工业品电商"
    ],
    # 制造业领域（全覆盖你的机器人/半导体/低空经济/汽车/军工相关文件名）
    "制造业": [
        "制造", "半导体", "机器人", "低空经济", "工业", "汽车", "军工", "3D打印",
        "无人机", "家电", "航天", "设备", "半导体材料", "半导体设备", "光刻机",
        "人形机器人", "工业机器人", "移动充电机器人", "智能机器人", "协作机器人",
        "商用服务机器人", "智能驾驶", "自动驾驶", "Robotaxi", "汽车后市场",
        "汽车玻璃", "智能网联", "军工贸", "国防军工", "高端装备", "海洋产业",
        "散热材料", "机械行业", "特种设备", "分布式存储", "算力中心",
        "商业航天", "低空经济发展", "家居后市场", "制造业出海"
    ]
}

# 2. 待分类文件类型（仅处理这些格式，后续新增格式加这里）
SUPPORTED_FILE_TYPES = [".pdf"]

# 3. 日志配置（是否输出详细分类日志）
PRINT_DETAIL_LOG = True


# ===================== 【工具函数区】- 核心逻辑，无需修改 =====================
def get_unique_filename(dst_path):
    """处理同名文件：自动加时间戳后缀，避免覆盖"""
    if not os.path.exists(dst_path):
        return dst_path
    # 拆分路径、文件名、后缀
    dst_dir = os.path.dirname(dst_path)
    file_name = os.path.basename(dst_path)
    name_without_ext, ext = os.path.splitext(file_name)
    # 生成时间戳后缀（精确到秒）
    timestamp = time.strftime("%Y%m%d%H%M%S", time.localtime())
    new_file_name = f"{name_without_ext}_{timestamp}{ext}"
    new_dst_path = os.path.join(dst_dir, new_file_name)
    return new_dst_path


def classify_filename(file_name):
    """
    核心分类函数：根据文件名匹配所属领域
    返回值：匹配到的领域名 / "未分类"
    """
    file_name_lower = file_name.lower()
    domain_count = defaultdict(int)

    # 统计各领域关键词出现次数（关键词匹配不区分大小写）
    for domain, keywords in DOMAIN_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in file_name_lower:
                domain_count[domain] += 1

    # 无匹配关键词 → 未分类
    if not domain_count or max(domain_count.values()) == 0:
        return "未分类"

    # 取关键词匹配次数最多的领域（次数相同取第一个）
    max_count = max(domain_count.values())
    matched_domains = [d for d, c in domain_count.items() if c == max_count]
    return matched_domains[0]


def validate_folder_path(folder_path):
    """校验文件夹路径合法性"""
    folder_path = folder_path.replace("\\", "/")
    if not os.path.exists(folder_path):
        return False, f"文件夹 {folder_path} 不存在！"
    if not os.path.isdir(folder_path):
        return False, f"{folder_path} 不是文件夹！"
    return True, folder_path


# ===================== 【主程序区】- 执行流程 =====================
def batch_classify_pdf_files():
    print("===== 【PDF文献自动分类工具】=====\n")

    # 1. 交互式获取待分类文件夹路径
    while True:
        folder_path = input("请输入PDF文件所在的文件夹完整路径：").strip()
        is_valid, msg = validate_folder_path(folder_path)
        if is_valid:
            folder_path = msg
            break
        print(f"❌ 错误：{msg} 请重新输入！")

    # 2. 扫描文件夹内所有支持的文件（仅PDF）
    target_files = []
    for file_name in os.listdir(folder_path):
        file_ext = os.path.splitext(file_name)[1].lower()
        if file_ext in SUPPORTED_FILE_TYPES:
            target_files.append(file_name)

    if not target_files:
        print("❌ 错误：该文件夹下无PDF文件！程序退出")
        return
    print(f"✅ 成功扫描到 {len(target_files)} 个PDF文件，开始分类...\n")

    # 3. 自动创建分类子文件夹（含"未分类"）
    domain_folders = {}
    all_domains = list(DOMAIN_KEYWORDS.keys()) + ["未分类"]
    for domain in all_domains:
        domain_folder = os.path.join(folder_path, domain)
        os.makedirs(domain_folder, exist_ok=True)
        domain_folders[domain] = domain_folder
        print(f"📂 已创建/确认文件夹：{domain_folder}")

    # 4. 批量分类+复制文件（核心步骤）
    classification_stats = defaultdict(int)  # 分类统计
    fail_files = []  # 失败文件列表

    for file_name in target_files:
        # 步骤4.1：确定文件所属领域
        domain = classify_filename(file_name)
        classification_stats[domain] += 1

        # 步骤4.2：构建源路径和目标路径
        src_path = os.path.join(folder_path, file_name)
        dst_path = os.path.join(domain_folders[domain], file_name)
        # 处理同名文件，避免覆盖
        dst_path = get_unique_filename(dst_path)

        # 步骤4.3：复制文件（保留原文件，仅复制）
        try:
            # copy2保留文件元数据（创建时间、修改时间等）
            shutil.copy2(src_path, dst_path)
            if PRINT_DETAIL_LOG:
                print(f"✅ 分类成功：{file_name} → {domain}")
        except Exception as e:
            fail_files.append((file_name, str(e)))
            print(f"⚠️  分类失败：{file_name} → 原因：{str(e)}")

    # 5. 输出分类结果汇总
    print("\n" + "=" * 50)
    print("📊 分类结果汇总")
    print("=" * 50)
    total_files = len(target_files)
    success_files = total_files - len(fail_files)
    print(f"总文件数：{total_files} | 成功分类：{success_files} | 失败：{len(fail_files)}")
    print("\n各领域分类数量：")
    for domain, count in sorted(classification_stats.items()):
        percentage = (count / total_files) * 100
        print(f"- {domain}：{count} 个（{percentage:.1f}%）")

    # 6. 输出失败文件（如有）
    if fail_files:
        print("\n❌ 失败文件列表：")
        for file_name, error in fail_files:
            print(f"- {file_name}：{error}")

    # 7. 输出结果文件位置
    print(f"\n🎉 分类完成！所有分类后的文件已保存至：{folder_path} 下的对应子文件夹")


# ===================== 【执行入口】 =====================
if __name__ == "__main__":
    # 程序启动提示
    print("欢迎使用PDF文献自动分类工具！")
    print("👉 特点：1. 仅复制文件（原文件保留）；2. 同名文件自动加时间戳；3. 支持扩展领域/关键词\n")
    # 执行主程序
    batch_classify_pdf_files()
    # 程序结束等待（避免控制台闪退）
    input("\n按回车键退出程序...")