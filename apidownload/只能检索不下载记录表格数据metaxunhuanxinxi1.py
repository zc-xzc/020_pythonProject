import requests
import os
import time
import csv

# ==================== 🔑 配置区 ====================
# 您的 3 把 Meta Key
API_KEY_POOL = [
    '3f7452135795adfb6bad21ee0d0613ad',
    '1784f9a81f12bcec431953bedcbf37fb',
    'd5735d354312e9c8f6e96c10a0ff9178',
]

BASE_URL = 'http://api.springernature.com/meta/v2/json'

# 关键词 (API 层面只搜词，本地再筛选期刊)
CORE_KEYWORD = "Environmental Regulation"
QUERY = CORE_KEYWORD

# 每日任务量
DAILY_TARGET = 300

# 文件路径
SAVE_DIR = 'papers_pdf_data'  # PDF保存位置
HISTORY_FILE = 'history_record.txt'  # 进度记录
METADATA_FILE = 'all_papers_info.csv'  # <--- 数据总表


# ===================================================

class KeyManager:
    def __init__(self, keys):
        self.keys = keys
        self.idx = 0
        self.banned = set()

    def get_key(self):
        if len(self.banned) >= len(self.keys):
            print("⚠️ 所有 Key 暂时休息，60秒后重试...")
            time.sleep(60)
            self.banned.clear()
        key = self.keys[self.idx]
        while key in self.banned:
            self.rotate()
            key = self.keys[self.idx]
        return key

    def rotate(self):
        self.idx = (self.idx + 1) % len(self.keys)

    def mark_bad(self, key, code):
        print(f"🚫 Key [...{key[-4:]}] 报错 {code}，暂时跳过")
        self.banned.add(key)
        self.rotate()


def load_start_index():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r') as f:
            try:
                return max(1, int(f.read().strip()))
            except:
                return 1
    return 1


def save_current_index(index):
    with open(HISTORY_FILE, 'w') as f:
        f.write(str(index))


# === 初始化 CSV 表格 ===
def init_csv_file():
    # 如果文件不存在，创建并写表头
    if not os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            # 这一行定义了表格里会有哪些列
            writer.writerow(['Title', 'DOI', 'Journal', 'Date', 'Authors', 'PDF_Status', 'Abstract'])


# === 保存一行数据 ===
def save_metadata_row(info):
    with open(METADATA_FILE, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow([
            info.get('title', ''),
            info.get('doi', ''),
            info.get('journal', ''),
            info.get('date', ''),
            info.get('authors', ''),
            info.get('pdf_status', ''),  # 下载状态
            info.get('abstract', '')  # 摘要放最后，因为比较长
        ])


def run_downloader():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    # 1. 启动前先准备好表格
    init_csv_file()

    key_mgr = KeyManager(API_KEY_POOL)
    start_index = load_start_index()
    end_index = start_index + DAILY_TARGET

    print(f"========================================")
    print(f"📊 Springer 全量数据记录器")
    print(f"📋 数据表: {METADATA_FILE} (无论下载是否成功都会记录)")
    print(f"🎯 目标: {start_index} -> {end_index - 1}")
    print(f"========================================\n")

    current_pointer = start_index
    batch_size = 10

    while current_pointer < end_index:
        actual_p = min(batch_size, end_index - current_pointer)

        try:
            current_key = key_mgr.get_key()
        except:
            break

        print(f"🔄 [请求] {current_pointer}-{current_pointer + actual_p} ...", end="")

        params = {
            'q': QUERY,
            'api_key': current_key,
            'p': actual_p,
            's': current_pointer
        }

        try:
            # 请求元数据
            response = requests.get(BASE_URL, params=params)

            if response.status_code != 200:
                print(f"\n❌ API 失败: {response.status_code}")
                key_mgr.mark_bad(current_key, response.status_code)
                time.sleep(1)
                continue

            data = response.json()
            records = data.get('records', [])

            if not records:
                print("\n⚠️  数据搜完了。")
                break

            print(f" ✅ API返回 {len(records)} 条")

            dl_headers = {'User-Agent': 'Mozilla/5.0'}

            for i, item in enumerate(records):
                # === 1. 严格筛选：只记录英文期刊 ===
                pub_type = item.get('publicationType', '')
                lang = item.get('language', '')

                # 如果不是期刊，或者不是英文，直接跳过，不记录也不下载
                if pub_type != 'Journal' or (lang and lang != 'en'):
                    continue

                # === 2. 提取信息 ===
                title = item.get('title', 'untitled').replace('/', '-').replace(':', '').strip()
                doi = item.get('doi')
                if not doi: continue

                creators = item.get('creators', [])
                authors = ", ".join([c.get('creator', '') for c in creators])

                # 准备数据包 (默认状态为 Unknown)
                article_info = {
                    'title': title,
                    'doi': doi,
                    'journal': item.get('publicationName', ''),
                    'date': item.get('publicationDate', ''),
                    'authors': authors,
                    'abstract': item.get('abstract', ''),
                    'pdf_status': 'Unknown'
                }

                # === 3. 尝试下载 PDF ===
                filename = f"{title[:150]}.pdf"
                filepath = os.path.join(SAVE_DIR, filename)

                print(f"   [{current_pointer + i}] 处理: {title[:15]}...", end="")

                if os.path.exists(filepath):
                    print(" [已存在]")
                    article_info['pdf_status'] = 'Existing'
                else:
                    pdf_url = f"https://link.springer.com/content/pdf/{doi}.pdf"
                    try:
                        r = requests.get(pdf_url, headers=dl_headers, stream=True, timeout=25)

                        # 判断是否为 PDF
                        content_type = r.headers.get('Content-Type', '')
                        if 'application/pdf' in content_type:
                            with open(filepath, 'wb') as f:
                                for chunk in r.iter_content(chunk_size=4096):
                                    f.write(chunk)
                            print(" [√ 下载成功]")
                            article_info['pdf_status'] = 'Downloaded'
                        else:
                            # 下载到了网页，说明要付费，或者没权限
                            print(" [X 无权限/付费]")
                            article_info['pdf_status'] = 'No Access (Paywall)'

                    except Exception as e:
                        print(" [X 网络错]")
                        article_info['pdf_status'] = 'Network Error'

                    time.sleep(1)  # 礼貌延时

                # === 4. 【关键】无论上面结果如何，强制保存元数据到表格 ===
                save_metadata_row(article_info)

            # 翻页逻辑
            key_mgr.rotate()
            current_pointer += actual_p
            save_current_index(current_pointer)
            time.sleep(2)

        except Exception as e:
            print(f"\n❌ 脚本异常: {e}")
            break

    print(f"\n🎉 任务完成！")
    print(f"📊 请查看表格文件: {METADATA_FILE}")


if __name__ == '__main__':
    run_downloader()