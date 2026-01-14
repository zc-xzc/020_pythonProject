import requests
import os
import time

# ==================== 🔑 配置区 ====================
API_KEY_POOL = [
    '7fd588e0e733f395811c51faec24c5cd',
    '0a0368743104858852a865ced7e5f176',
    '68dcdb2552f4b6b447f886735892aa22',
]

BASE_URL = 'http://api.springernature.com/meta/v2/json'

# 🔥 核心修改：只搜 SCI期刊(Journal) + 英文(en) + 开放获取(OA)
CORE_KEYWORD = "Environmental Regulation"
# 这里的 openaccess:true 是关键，它保证了下载必成功
QUERY = f'keyword:"{CORE_KEYWORD}" AND openaccess:true'

DAILY_TARGET = 300
SAVE_DIR = 'papers_sci_free_oa'  # 换个文件夹存，区分一下
HISTORY_FILE = 'history_oa_only.txt'


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


def run_downloader():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    key_mgr = KeyManager(API_KEY_POOL)
    start_index = load_start_index()
    end_index = start_index + DAILY_TARGET

    print(f"========================================")
    print(f"🟢 Springer 必成版 (只下免费SCI)")
    print(f"🔑 Key 数量: {len(API_KEY_POOL)}")
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
            'q': QUERY,  # 这里的 QUERY 已经过滤了付费文章
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

            # 下载流程
            dl_headers = {'User-Agent': 'Mozilla/5.0'}
            valid_count = 0

            for i, item in enumerate(records):
                # 本地再次筛选确保是期刊+英文
                pub_type = item.get('publicationType', '')
                lang = item.get('language', '')

                if pub_type != 'Journal' or (lang and lang != 'en'):
                    continue

                valid_count += 1
                title = item.get('title', 'untitled').replace('/', '-').replace(':', '')
                doi = item.get('doi')

                if not doi: continue

                filename = f"{title[:150]}.pdf"
                filepath = os.path.join(SAVE_DIR, filename)

                if os.path.exists(filepath):
                    print(f"   [跳过] 已存在: {title[:10]}...")
                    continue

                pdf_url = f"https://link.springer.com/content/pdf/{doi}.pdf"

                try:
                    print(f"   [{current_pointer + i}] 下载(OA): {title[:15]}...", end="")
                    r = requests.get(pdf_url, headers=dl_headers, stream=True, timeout=30)

                    if 'application/pdf' in r.headers.get('Content-Type', ''):
                        with open(filepath, 'wb') as f:
                            for chunk in r.iter_content(chunk_size=4096):
                                f.write(chunk)
                        print(" [√ 成功]")
                    else:
                        print(" [X 异常]")  # 理论上加了 openaccess:true 不会出现这情况
                except:
                    print(" [X 网络错]")
                time.sleep(1)

            key_mgr.rotate()
            current_pointer += actual_p
            save_current_index(current_pointer)
            time.sleep(2)

        except Exception as e:
            print(f"❌ 异常: {e}")
            break

    print(f"\n🎉 任务完成！下次从 {load_start_index()} 开始。")


if __name__ == '__main__':
    run_downloader()