import requests
import os
import time

# === 再次尝试 Meta API (通常数据更全) ===
API_KEY = 'd5735d354312e9c8f6e96c10a0ff9178'  # Meta API Key
BASE_URL = 'http://api.springernature.com/meta/v2/json'
# ==========================================

QUERY = 'Environmental Regulation'
SAVE_DIR = 'papers_pdf_meta'
MAX_PAPERS = 5


def try_meta_download():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    print(f"正在重试 Meta API 下载流程...")

    # 尝试获取数据
    params = {'q': QUERY, 'api_key': API_KEY, 'p': MAX_PAPERS, 's': 1}

    try:
        response = requests.get(BASE_URL, params=params)

        if response.status_code == 401:
            print(">>> 依然 401: 请务必去后台把 'Website URL' 改为 'http://localhost'")
            return

        response.raise_for_status()
        data = response.json()

        print(f"Meta API 连接成功！准备下载 {len(data.get('records', []))} 个文件。\n")

        for i, item in enumerate(data.get('records', [])):
            title = item.get('title', 'untitled').replace('/', '-').replace(':', '-')

            # Meta API 的标准 PDF 查找逻辑
            pdf_url = None
            for u in item.get('url', []):
                if u.get('format') == 'pdf':
                    pdf_url = u.get('value')
                    break

            if pdf_url:
                print(f"[{i + 1}] 正在下载: {title[:20]}...")
                # 真实下载
                r = requests.get(pdf_url)
                with open(os.path.join(SAVE_DIR, f"{title}.pdf"), 'wb') as f:
                    f.write(r.content)
            else:
                print(f"[{i + 1}] 跳过: 无 PDF 链接 (可能是图书章节)")

    except Exception as e:
        print(f"出错: {e}")


if __name__ == '__main__':
    try_meta_download()