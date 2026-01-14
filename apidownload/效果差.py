import requests
import os
import time

# ==================== 🔑 配置区 ====================
API_KEY_POOL = [
    '3f7452135795adfb6bad21ee0d0613ad',
    '1784f9a81f12bcec431953bedcbf37fb',
    'd5735d354312e9c8f6e96c10a0ff9178',
]

# Meta 接口
BASE_URL = 'http://api.springernature.com/meta/v2/json'

# 【修改点1】只搜关键词，不加高级筛选，避开 403 报错
CORE_KEYWORD = "Environmental Regulation"
QUERY = CORE_KEYWORD

# 每日任务
DAILY_TARGET = 300
SAVE_DIR = 'papers_sci_english'
HISTORY_FILE = 'sci_download_history.txt'


# ===================================================

class KeyManager:
    def __init__(self, keys):
        self.keys = keys
        self.idx = 0
        self.banned = set()

    def get_key(self):
        # 简单轮询
        if len(self.banned) >= len(self.keys):
            # 如果全封了，强制重置（死马当活马医，防止程序直接退出）
            print("⚠️ 所有 Key 暂时不可用，休息 60 秒后重试...")
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
        print(f"🚫 Key [...{key[-4:]}] 报错 {code}，暂时停用")
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


def run_downloader():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    key_mgr = KeyManager(API_KEY_POOL)
    start_index = load_start_index()
    end_index = start_index + DAILY_TARGET

    print(f"========================================")
    print(f"🎓 Springer SCI 下载器 (本地筛选版)")
    print(f"🔑 Key 数量: {len(API_KEY_POOL)}")
    print(f"🎯 目标: {start_index} -> {end_index - 1}")
    print(f"🛡️ 策略: API只搜词 -> 本地筛期刊/英文")
    print(f"========================================\n")

    current_pointer = start_index
    # 【修改点2】把每批次请求降低到 10，这是 Meta API 的安全线
    batch_size = 10

    while current_pointer < end_index:
        actual_p = min(batch_size, end_index - current_pointer)

        try:
            current_key = key_mgr.get_key()
        except:
            break

        print(f"🔄 [请求] {current_pointer}-{current_pointer + actual_p} (Key:..{current_key[-4:]}) ...")

        params = {
            'q': QUERY,
            'api_key': current_key,
            'p': actual_p,
            's': current_pointer
        }

        try:
            # 这里的请求必须非常纯净，不带 Headers
            response = requests.get(BASE_URL, params=params)

            if response.status_code != 200:
                print(f"❌ API 失败: {response.status_code} | {response.text[:100]}")
                key_mgr.mark_bad(current_key, response.status_code)
                time.sleep(2)
                continue

            data = response.json()
            records = data.get('records', [])

            if not records:
                print("⚠️  数据搜完了。")
                break

            print(f"   ✅ API 返回 {len(records)} 条，正在本地筛选...")

            # === 下载 + 本地筛选流程 ===
            dl_headers = {'User-Agent': 'Mozilla/5.0'}

            valid_count = 0
            for i, item in enumerate(records):
                title = item.get('title', 'untitled').replace('/', '-').replace(':', '')
                doi = item.get('doi')

                # 【修改点3】核心：本地筛选逻辑
                # 1. 检查是不是期刊 (publicationType == 'Journal')
                pub_type = item.get('publicationType', '')
                # 2. 检查是不是英文 (language == 'en')
                lang = item.get('language', '')

                # 打印一下类型，方便您观察
                # print(f"      (分析: 类型={pub_type}, 语言={lang})")

                if pub_type != 'Journal':
                    # 如果是书、会议，就跳过
                    continue
                if lang and lang != 'en':
                    # 如果明确标了不是英文，跳过
                    continue

                valid_count += 1

                if not doi: continue

                filename = f"{title[:150]}.pdf"
                filepath = os.path.join(SAVE_DIR, filename)

                if os.path.exists(filepath):
                    print(f"   [跳过] 已存在: {title[:10]}...")
                    continue

                pdf_url = f"https://link.springer.com/content/pdf/{doi}.pdf"

                try:
                    print(f"   [{current_pointer + i}] 下载(SCI): {title[:15]}...", end="")
                    r = requests.get(pdf_url, headers=dl_headers, stream=True, timeout=30)

                    if 'application/pdf' in r.headers.get('Content-Type', ''):
                        with open(filepath, 'wb') as f:
                            for chunk in r.iter_content(chunk_size=4096):
                                f.write(chunk)
                        print(" [√]")
                    else:
                        print(" [X 非PDF/付费]")
                except:
                    print(" [X 网络错]")
                time.sleep(1)

            print(f"   📊 本批次 {len(records)} 条中，符合筛选条件的有 {valid_count} 条。")

            # 无论这批次里有几个符合条件的，都要往下翻页
            key_mgr.rotate()
            current_pointer += actual_p
            save_current_index(current_pointer)
            time.sleep(2)

        except Exception as e:
            print(f"❌ 脚本异常: {e}")
            break

    print(f"\n🎉 任务完成！下次从 {load_start_index()} 开始。")


if __name__ == '__main__':
    run_downloader()