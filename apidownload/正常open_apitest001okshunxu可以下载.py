import requests
import os
import time

# ================= 配置区 =================
# 1. Key (保持不变)
API_KEY = '68dcdb2552f4b6b447f886735892aa22'
BASE_URL = 'http://api.springernature.com/openaccess/json'

# 2. 搜索设置
QUERY = 'pollution discharge'
SAVE_DIR = 'papers_database'

# 3. 任务量设置
DAILY_TARGET = 50  # 每天总共要下载 50 篇
API_PAGE_SIZE = 10  # <---【核心修改】每次只向服务器要 10 条，避免触发“高级功能”报错

# 4. 历史记录文件
HISTORY_FILE = 'download_history.txt'


# =========================================

def load_start_index():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r') as f:
            try:
                val = int(f.read().strip())
                return max(1, val)
            except:
                return 1
    return 1


def save_current_index(index):
    with open(HISTORY_FILE, 'w') as f:
        f.write(str(index))


def run_task():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    # 1. 规划任务
    start_index = load_start_index()
    end_index = start_index + DAILY_TARGET

    print(f"========================================")
    print(f"📖 进度读取: 从第 {start_index} 篇开始")
    print(f"🎯 今日目标: 下载到第 {end_index - 1} 篇 (共{DAILY_TARGET}篇)")
    print(f"⚡ 策略: 分批请求，每批 {API_PAGE_SIZE} 条")
    print(f"========================================\n")

    current_pointer = start_index

    while current_pointer < end_index:
        # 即使还剩很多，每次也最多只取 10 条
        actual_p = min(API_PAGE_SIZE, end_index - current_pointer)

        print(f"🔄 [API请求] 第 {current_pointer} - {current_pointer + actual_p} 条 ...")

        params = {
            'q': QUERY,
            'api_key': API_KEY,
            'p': actual_p,  # 这里现在是 10，服务器会很乐意接受
            's': current_pointer
        }

        try:
            # 1. 获取列表 (不带 Headers，保持纯净)
            response = requests.get(BASE_URL, params=params)

            if response.status_code != 200:
                print(f"❌ API 报错: {response.status_code}")
                print(f"   原因: {response.text}")
                # 如果还是不行，可能需要更小的粒度，或者暂停一下
                break

            data = response.json()
            records = data.get('records', [])

            if not records:
                print("⚠️  没有更多数据了，任务提前结束。")
                break

            print(f"   ✅ 成功获取 {len(records)} 条元数据，开始下载...")

            # 2. 遍历下载 (带 Headers 伪装)
            download_headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }

            for i, item in enumerate(records):
                title = item.get('title', 'untitled').replace('/', '-').replace(':', '-').replace('?', '')
                doi = item.get('doi')

                # 跳过无 DOI
                if not doi:
                    continue

                filepath = os.path.join(SAVE_DIR, f"{title}.pdf")
                if os.path.exists(filepath):
                    print(f"   [已存在] 跳过: {title[:15]}...")
                    continue

                # 构造下载链接
                pdf_url = f"https://link.springer.com/content/pdf/{doi}.pdf"

                try:
                    print(f"   [{current_pointer + i}] 下载中: {title[:20]}...", end="")
                    r = requests.get(pdf_url, headers=download_headers, stream=True, timeout=30)

                    if 'application/pdf' in r.headers.get('Content-Type', ''):
                        with open(filepath, 'wb') as f:
                            for chunk in r.iter_content(chunk_size=1024):
                                if chunk:
                                    f.write(chunk)
                        print(" [√]")
                    else:
                        print(" [X 非PDF]")

                except Exception as e:
                    print(f" [X 网络错: {e}]")

                # 稍微休息一下，别太快
                time.sleep(1)

            # 更新大循环指针
            current_pointer += actual_p
            save_current_index(current_pointer)

            # 批次之间休息 2 秒，模拟人类翻页
            time.sleep(2)

        except Exception as e:
            print(f"❌ 程序发生错误: {e}")
            break

    print(f"\n🎉 今日任务结束！进度已保存至第 {load_start_index()} 篇。")


if __name__ == '__main__':
    run_task()