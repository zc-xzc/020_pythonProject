# -*- coding: utf-8 -*-
"""
批量下载论文工具（2025交互式最终版）
特性：
1. 命令行交互式配置（代理/路径/搜索模式），无需改代码
2. 内置15+最新Sci-Hub镜像，自动重试
3. 失败论文自动导出CSV清单，方便手动处理
4. 分块下载大文件、自动修复URL拼接错误
5. 兼容多编码Bib文件、生成安全文件名
"""
import time
import re
import requests
from lxml import etree
import sys
import os
import csv

# 禁用requests警告（HTTPS证书问题）
requests.packages.urllib3.disable_warnings()

# 2025年12月全网最全可用Sci-Hub镜像列表（按优先级排序）
SCI_HUB_MIRRORS = [
    'https://sci-hub.se/',
    'https://sci-hub.st/',
    'https://sci-hub.ru/',
    'https://sci-hub.ee/',
    'https://sci-hub.cc/',
    'https://sci-hub.wf/',
    'https://sci-hub.ren/',
    'https://sci-hub.mosf.ru/',
    'https://sci-hub.is/',
    'https://sci-hub.es/',
    'https://sci-hub.gg/',
    'https://sci-hub.tw/',
    'https://sci-hub.hk/',
    'https://sci-hub.la/',
    'https://sci-hub.su/',
    'https://sci-hub.ac/'
]


def get_interactive_config():
    """
    交互式获取命令行配置（核心：所有配置都从命令行输入）
    :return: (PROXIES, BIB_FILE_PATH, SAVE_FOLDER, SEARCH_MODE)
    """
    print("=" * 50)
    print("📋 批量论文下载工具 - 配置向导")
    print("=" * 50)

    # ---------------------- 1. 代理配置 ----------------------
    print("\n【1/4 代理配置】")
    use_proxy = input("是否使用代理（解决国内IP限制）？(y/n，默认n)：").strip().lower()
    PROXIES = None
    if use_proxy == 'y':
        proxy_addr = input("请输入代理地址（格式：127.0.0.1:10809）：").strip()
        if proxy_addr:
            PROXIES = {
                'http': f'http://{proxy_addr}',
                'https': f'http://{proxy_addr}'
            }
            print(f"✅ 代理已配置：{PROXIES}")
        else:
            print("⚠️  代理地址为空，将不使用代理")
    else:
        print("✅ 不使用代理")

    # ---------------------- 2. Bib文件路径配置 ----------------------
    print("\n【2/4 Bib文件路径配置】")
    default_bib_path = r"D:\Hefei_University_of_Technology_Work\020_pythonProject\apidownload\savedrecs.bib"
    bib_path = input(f"请输入Bib文件路径（默认：{default_bib_path}）：").strip()
    BIB_FILE_PATH = bib_path if bib_path else default_bib_path
    # 校验Bib文件是否存在
    while not os.path.exists(BIB_FILE_PATH):
        print(f"❌ 找不到文件：{BIB_FILE_PATH}")
        bib_path = input("请重新输入正确的Bib文件路径：").strip()
        BIB_FILE_PATH = bib_path if bib_path else default_bib_path
    print(f"✅ Bib文件路径：{BIB_FILE_PATH}")

    # ---------------------- 3. PDF保存路径配置 ----------------------
    print("\n【3/4 PDF保存路径配置】")
    default_save_path = r"D:\Hefei_University_of_Technology_Work\020_pythonProject\apidownload\paper"
    save_path = input(f"请输入PDF保存路径（默认：{default_save_path}）：").strip()
    SAVE_FOLDER = save_path if save_path else default_save_path
    # 自动创建保存文件夹
    os.makedirs(SAVE_FOLDER, exist_ok=True)
    print(f"✅ PDF保存路径：{SAVE_FOLDER}（文件夹不存在已自动创建）")

    # ---------------------- 4. 搜索模式配置 ----------------------
    print("\n【4/4 搜索模式配置】")
    SEARCH_MODE = None
    while SEARCH_MODE not in [2, 3]:
        mode_input = input("请选择搜索模式（2=按DOI搜索（推荐），3=按标题搜索，默认2）：").strip()
        if not mode_input:
            SEARCH_MODE = 2
        else:
            try:
                SEARCH_MODE = int(mode_input)
                if SEARCH_MODE not in [2, 3]:
                    print("❌ 仅支持2或3，请重新输入")
            except ValueError:
                print("❌ 请输入数字2或3")
    print(f"✅ 搜索模式：{'按DOI搜索' if SEARCH_MODE == 2 else '按标题搜索'}")

    print("\n" + "=" * 50)
    print("📌 所有配置已完成，即将开始处理...")
    print("=" * 50)
    time.sleep(1)  # 暂停1秒，让用户确认配置

    return PROXIES, BIB_FILE_PATH, SAVE_FOLDER, SEARCH_MODE


def search_paper(artName, current_base_url, proxies):
    """
    单镜像搜索逻辑（抽离为子函数，方便复用）
    :param artName: 清理后的DOI/标题
    :param current_base_url: 当前使用的Sci-Hub镜像
    :param proxies: 代理配置
    :return: 完整PDF下载链接 | None
    """
    try:
        url = f"{current_base_url}{artName}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        }

        # 发送请求（支持代理）
        res = requests.get(
            url,
            headers=headers,
            proxies=proxies,
            timeout=10,
            allow_redirects=True,
            verify=False
        )
        res.raise_for_status()
        tree = etree.HTML(res.text)

        # 方案1：iframe#pdf的src（主流结构）
        pdf_url = tree.xpath('//iframe[@id="pdf"]/@src')
        if pdf_url and pdf_url[0]:
            raw_url = pdf_url[0].strip()
            # 修复1：处理各种相对路径场景
            if raw_url.startswith('//'):
                full_url = f"https:{raw_url}"
            elif raw_url.startswith('/'):
                # 关键：拼接当前镜像的域名（而非固定值）
                full_url = f"{current_base_url.rstrip('/')}{raw_url}"
            elif raw_url.startswith('http'):
                full_url = raw_url
            else:
                full_url = f"{current_base_url.rstrip('/')}/{raw_url.lstrip('/')}"

            # 修复2：补充缺失的.pdf后缀（部分链接截断）
            if not full_url.endswith('.pdf') and not '?download' in full_url:
                full_url = f"{full_url}.pdf"

            # 验证URL格式
            if full_url.startswith(('http://', 'https://')):
                return full_url

        # 方案2：直接的下载按钮链接
        pdf_url = tree.xpath('//a[contains(text(), "download") or contains(@href, ".pdf")]/@href')
        if pdf_url and pdf_url[0]:
            raw_url = pdf_url[0].strip()
            if raw_url.startswith('//'):
                return f"https:{raw_url}"
            elif raw_url.startswith('/'):
                return f"{current_base_url.rstrip('/')}{raw_url}"
            elif raw_url.startswith('http'):
                return raw_url
            else:
                return f"{current_base_url}{raw_url}"

        # 方案3：尝试从script中提取PDF链接
        script_text = tree.xpath('//script[contains(text(), "pdf")]/text()')
        for script in script_text:
            pdf_match = re.search(r'https?://[^\s"]+\.pdf', script)
            if pdf_match:
                return pdf_match.group()

    except Exception as e:
        # 仅打印关键错误，避免日志冗余
        if "404" not in str(e) and "Forbidden" not in str(e):
            print(f"⚠️  镜像 {current_base_url} 访问失败：{str(e)[:50]}")
    return None


def batch_search_paper(artName, proxies):
    """
    多镜像批量搜索（主函数）
    :param artName: DOI/标题
    :param proxies: 代理配置
    :return: 有效PDF链接 | None
    """
    # 清理关键词
    artName = re.sub(r'[\s\t\n\r]', '', artName.strip())
    if not artName:
        return None

    # 遍历镜像重试
    for base_url in SCI_HUB_MIRRORS:
        pdf_url = search_paper(artName, base_url, proxies)
        if pdf_url:
            return pdf_url

    return None


def download_paper(downUrl_in, proxies):
    """
    下载PDF文件（增加重试+分块下载）
    :param downUrl_in: PDF下载链接
    :param proxies: 代理配置
    :return: PDF二进制内容 | None
    """
    if not downUrl_in:
        return None

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/pdf,application/x-pdf;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Connection': 'keep-alive',
        'Range': 'bytes=0-'  # 支持断点续传
    }

    # 最多重试2次
    for retry in range(2):
        try:
            res = requests.get(
                downUrl_in,
                headers=headers,
                proxies=proxies,
                timeout=60,  # 延长超时（大文件）
                stream=True,
                verify=False
            )
            res.raise_for_status()

            # 分块下载（避免内存溢出）
            pdf_content = b''
            for chunk in res.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    pdf_content += chunk
            return pdf_content

        except Exception as e:
            print(f"❌ 下载重试{retry + 1}失败：{str(e)[:50]}")
            time.sleep(1)
            continue

    return None


def into_bib(file_tix_in):
    """
    解析BibTex文件，提取作者/年份/DOI/标题（兼容多种格式）
    :param file_tix_in: Bib文件路径
    :return: [作者列表, 年份列表, DOI列表, 标题列表]
    """
    if not os.path.exists(file_tix_in):
        print(f"❌ 找不到Bib文件：{file_tix_in}")
        return [[], [], [], []]

    # 兼容utf-8/gbk编码
    try:
        with open(file_tix_in, mode='r', encoding='utf-8') as file:
            content = file.read()
    except UnicodeDecodeError:
        with open(file_tix_in, mode='r', encoding='gbk') as file:
            content = file.read()

    # 优化正则：兼容大小写/空格/单双花括号/跨行内容
    pattern_author = re.compile(r'author\s*=\s*{(.*?)}', re.IGNORECASE | re.DOTALL)
    pattern_year = re.compile(r'year\s*=\s*{(.*?)}', re.IGNORECASE)
    pattern_doi = re.compile(r'doi\s*=\s*{(.*?)}', re.IGNORECASE)
    pattern_title = re.compile(r'title\s*=\s*{(.*?)}', re.IGNORECASE | re.DOTALL)

    # 提取并清理字段
    match_author = [re.sub(r'[{}|\n|\t]', '', m.strip()) for m in pattern_author.findall(content)]
    match_year = [re.sub(r'[{}|\n|\t]', '', m.strip()) for m in pattern_year.findall(content)]
    match_doi = [re.sub(r'[{}|\n|\t|\s]', '', m.strip()) for m in pattern_doi.findall(content)]
    match_title = [re.sub(r'[{}|\n|\t]', '', m.strip()) for m in pattern_title.findall(content)]

    # 调试信息
    print(f"\n📊 Bib文件解析结果：")
    print(f"   作者数：{len(match_author)} | 年份数：{len(match_year)}")
    print(f"   DOI数：{len(match_doi)} | 标题数：{len(match_title)}")
    if match_doi:
        print(f"   前3个DOI示例：{match_doi[:3]}")

    return [match_author, match_year, match_doi, match_title]


def export_fail_list(fail_list, doi_list, save_folder):
    """
    导出失败论文列表到CSV文件
    :param fail_list: 失败论文[(编号, 标题), ...]
    :param doi_list: 所有DOI列表
    :param save_folder: 保存路径
    """
    fail_csv_path = os.path.join(save_folder, "失败论文列表.csv")
    try:
        with open(fail_csv_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["论文编号", "DOI", "标题"])
            for num, title in fail_list:
                doi = doi_list[num - 1] if (num - 1) < len(doi_list) else ""
                writer.writerow([num, doi, title])
        print(f"\n📝 失败列表已导出至：{fail_csv_path}")
    except Exception as e:
        print(f"⚠️  失败列表导出失败：{str(e)}")


if __name__ == '__main__':
    # ===================== 交互式获取所有配置 =====================
    PROXIES, BIB_FILE_PATH, SAVE_FOLDER, SEARCH_MODE = get_interactive_config()

    # 解析Bib文件
    paper_info = into_bib(BIB_FILE_PATH)
    author_list, year_list, doi_list, title_list = paper_info
    total_papers = len(year_list)

    # 检查解析结果
    if total_papers == 0:
        print("\n❌ 未解析到任何论文！请检查Bib文件格式")
        sys.exit(1)
    print(f"\n🚀 开始处理 {total_papers} 篇论文...")
    print(f"🔌 代理配置：{'已启用' if PROXIES else '未启用'}")
    print(f"🌐 镜像数量：{len(SCI_HUB_MIRRORS)} 个")

    # 检查DOI完整性（仅DOI模式下）
    if SEARCH_MODE == 2 and len(doi_list) != total_papers:
        print(f"⚠️  警告：DOI数量({len(doi_list)})与论文数({total_papers})不匹配！")
        print("   建议：1. 补充Bib文件中的DOI  2. 重新运行选择标题搜索模式")

    # 批量下载
    fail_list = []
    success_list = []
    for idx in range(total_papers):
        print(f"\n--- 处理第 {idx + 1}/{total_papers} 篇 ---")

        # 获取搜索关键词
        if SEARCH_MODE == 2:
            search_key = doi_list[idx] if idx < len(doi_list) else ""
            print(f"DOI：{search_key}")
        else:
            search_key = title_list[idx] if idx < len(title_list) else ""
            print(f"标题：{search_key[:50]}...")

        # 跳过空关键词
        if not search_key:
            print("❌ 无搜索关键词，跳过")
            fail_list.append((idx + 1, title_list[idx] if idx < len(title_list) else "未知标题"))
            continue

        # 搜索PDF链接
        print("🔍 正在搜索PDF链接...")
        pdf_url = batch_search_paper(search_key, PROXIES)
        if not pdf_url:
            print("❌ 未找到PDF链接（所有镜像均失败）")
            fail_list.append((idx + 1, title_list[idx] if idx < len(title_list) else "未知标题"))
            continue
        print(f"🔗 找到PDF链接：{pdf_url[:80]}...")

        # 下载PDF
        print("📥 正在下载PDF...")
        pdf_content = download_paper(pdf_url, PROXIES)
        if not pdf_content:
            print("❌ PDF下载失败（重试后仍失败）")
            fail_list.append((idx + 1, title_list[idx] if idx < len(title_list) else "未知标题"))
            continue

        # 生成安全文件名
        author = author_list[idx] if idx < len(author_list) else "Unknown"
        year = year_list[idx] if idx < len(year_list) else "Unknown"
        # 过滤非法字符+截断过长作者名
        safe_author = re.sub(r'[\\/:*?"<>|,]', '_', author[:20])
        file_name = f"{safe_author}_{year}_{idx + 1}.pdf"
        save_path = os.path.join(SAVE_FOLDER, file_name)

        # 保存文件
        try:
            with open(save_path, 'wb') as f:
                f.write(pdf_content)
            print(f"✅ 保存成功：{file_name}")
            success_list.append((idx + 1, file_name))
        except Exception as e:
            print(f"❌ 保存失败：{str(e)}")
            fail_list.append((idx + 1, title_list[idx] if idx < len(title_list) else "未知标题"))

        # 延长延迟（避免反爬）
        time.sleep(2)

    # 最终统计
    print(f"\n========== 任务完成 ==========")
    print(f"✅ 成功下载：{len(success_list)} 篇")
    print(f"❌ 下载失败：{len(fail_list)} 篇")

    if success_list:
        print("\n✅ 成功列表（前10条）：")
        for num, fname in success_list[:10]:
            print(f"   第{num}篇：{fname}")

    if fail_list:
        print("\n❌ 失败列表（前10条）：")
        for num, title in fail_list[:10]:
            print(f"   第{num}篇：{title[:60]}...")
        # 导出失败列表到CSV
        export_fail_list(fail_list, doi_list, SAVE_FOLDER)

    print(f"\n📁 PDF保存根路径：{SAVE_FOLDER}")
    print("\n💡 失败原因可能：1.Sci-Hub未收录 2.网络/地区限制 3.DOI无效")
    print("💡 提升成功率：启用代理 + 定期更新镜像列表")