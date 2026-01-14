import requests
import os
import time
import csv  # <--- 新增：用于处理表格保存

# ==================== 核心配置区 ====================
# 您的 Meta Key 池
API_KEY_POOL = [
    '3f7452135795adfb6bad21ee0d0613ad',
    '1784f9a81f12bcec431953bedcbf37fb',
    'd5735d354312e9c8f6e96c10a0ff9178',
]

BASE_URL = 'http://api.springernature.com/meta/v2/json'

# 搜索关键词 (只搜 SCI/英文期刊)
CORE_KEYWORD = "pollution discharge"
QUERY = CORE_KEYWORD

# 每日任务量
DAILY_TARGET = 300

# 保存路径
SAVE_DIR = 'papers_full_data'  # PDF 保存文件夹
HISTORY_FILE = 'history_full.txt'  # 进度记录
METADATA_FILE = 'papers_metadata.csv'  # <--- 新增：所有文章信息的总表


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


# === 新增功能：初始化 CSV 表格 ===
def init_csv_file():
    # 如果文件不存在，就创建并写入表头
    if not os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            # 定义表头：题目，DOI，期刊，日期，作者，摘要，PDF是否下载
            writer.writerow(['Title', 'DOI', 'Journal', 'Date', 'Authors', 'Abstract', 'PDF_Status'])


# === 新增功能：保存一行数据到 CSV ===
def save_metadata_row(info_dict):
    with open(METADATA_FILE, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow([
            info_dict.get('title', ''),
            info_dict.get('doi', ''),
            info_dict.get('journal', ''),
            info_dict.get('date', ''),
            info_dict.get('authors', ''),
            info_dict.get('abstract', ''),  # <--- 最重要的数据：摘要
            info_dict.get('pdf_status', '')
        ])


def run_downloader():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    # 初始化表格
    init_csv_file()

    key_mgr = KeyManager(API_KEY_POOL)
    start_index = load_start_index()
    end_index = start_index + DAILY_TARGET

    print(f"========================================")
    print(f"📊 Springer 全能收割机 (PDF + Excel)")
    print(f"📄 元数据表: {METADATA_FILE}")
    print(f"📂 PDF文件夹: {SAVE_DIR}")
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

        print(f"🔄 [请求API] {current_pointer}-{current_pointer + actual_p} ...", end="")

        params = {
            'q': QUERY,
            'api_key': current_key,
            'p': actual_p,
            's': current_pointer
        }

        try:
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

            print(f" ✅ 获取 {len(records)} 条")

            dl_headers = {'User-Agent': 'Mozilla/5.0'}

            for i, item in enumerate(records):
                # 1. 本地筛选 (只看英文期刊)
                pub_type = item.get('publicationType', '')
                lang = item.get('language', '')
                if pub_type != 'Journal' or (lang and lang != 'en'):
                    continue

                title = item.get('title', 'untitled').replace('/', '-').replace(':', '')
                doi = item.get('doi')
                if not doi: continue

                # 2. === 提取元数据 (核心修改) ===
                # 提取作者
                creators = item.get('creators', [])
                author_names = ", ".join([c.get('creator', '') for c in creators])

                # 准备要保存的数据包
                article_info = {
                    'title': title,
                    'doi': doi,
                    'journal': item.get('publicationName', ''),
                    'date': item.get('publicationDate', ''),
                    'authors': author_names,
                    'abstract': item.get('abstract', ''),  # 这就是您以后文本挖掘要用的东西
                    'pdf_status': 'Failed'  # 默认失败，下载成功后改写
                }

                # 3. 下载 PDF
                filename = f"{title[:150]}.pdf"
                filepath = os.path.join(SAVE_DIR, filename)

                # 如果文件已存在，标记为 Existing
                if os.path.exists(filepath):
                    print(f"   [跳过] 已存在: {title[:10]}...")
                    article_info['pdf_status'] = 'Existing'
                else:
                    # 尝试下载
                    pdf_url = f"https://link.springer.com/content/pdf/{doi}.pdf"
                    try:
                        print(f"   [{current_pointer + i}] 下载: {title[:15]}...", end="")
                        r = requests.get(pdf_url, headers=dl_headers, stream=True, timeout=30)

                        if 'application/pdf' in r.headers.get('Content-Type', ''):
                            with open(filepath, 'wb') as f:
                                for chunk in r.iter_content(chunk_size=4096):
                                    f.write(chunk)
                            print(" [√ 成功]")
                            article_info['pdf_status'] = 'Downloaded'
                        else:
                            print(" [X 非PDF]")
                            article_info['pdf_status'] = 'No Access'
                    except:
                        print(" [X 网络错]")
                        article_info['pdf_status'] = 'Error'
                    time.sleep(1)

                # 4. === 实时写入 CSV 表格 ===
                # 不管 PDF 下载成不成功，信息都得留下来！
                save_metadata_row(article_info)

            key_mgr.rotate()
            current_pointer += actual_p
            save_current_index(current_pointer)
            time.sleep(2)

        except Exception as e:
            print(f"❌ 异常: {e}")
            break

    print(f"\n🎉 任务完成！")
    print(f"📊 所有文章信息已保存在: {METADATA_FILE}")


if __name__ == '__main__':
    run_downloader()