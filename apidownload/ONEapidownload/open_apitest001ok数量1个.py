import requests
import os
import time

# ================= 配置区 =================
# 依然使用您当前能跑通的 Open Access Key
API_KEY = '68dcdb2552f4b6b447f886735892aa22'
BASE_URL = 'http://api.springernature.com/openaccess/json'

# 搜索关键词
QUERY = 'Environmental Regulation'
SAVE_DIR = 'papers_oa_fixed'
MAX_PAPERS = 5


# =========================================

def download_by_doi_construction():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)
        print(f"创建文件夹: {SAVE_DIR}")

    print(f"开始搜索: {QUERY} (使用 DOI 拼接法)...")

    params = {
        'q': QUERY,
        'api_key': API_KEY,
        'p': MAX_PAPERS,
        's': 1
    }

    try:
        # 1. 获取元数据
        response = requests.get(BASE_URL, params=params)
        response.raise_for_status()
        data = response.json()

        records = data.get('records', [])
        print(f"API 响应成功，获取到 {len(records)} 条记录。\n")

        for i, item in enumerate(records):
            title = item.get('title', 'untitled').replace('/', '-').replace(':', '-')
            doi = item.get('doi')  # 直接获取 DOI，例如 10.1007/s00467...

            if doi:
                # === 核心修改：暴力拼接 PDF 链接 ===
                # 规律：https://link.springer.com/content/pdf/{DOI}.pdf
                pdf_url = f"https://link.springer.com/content/pdf/{doi}.pdf"

                print(f"[{i + 1}] 正在下载: {title[:20]}...")
                print(f"    -> 构造链接: {pdf_url}")

                try:
                    # 模拟浏览器请求，防止被拦截
                    headers = {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                    }
                    r = requests.get(pdf_url, headers=headers, stream=True)

                    # 检查是不是下载到了真正的 PDF (而不是报错网页)
                    content_type = r.headers.get('Content-Type', '')
                    if 'application/pdf' in content_type:
                        file_path = os.path.join(SAVE_DIR, f"{title}.pdf")
                        with open(file_path, 'wb') as f:
                            for chunk in r.iter_content(chunk_size=1024):
                                if chunk:
                                    f.write(chunk)
                        print(f"    -> √ 下载成功")
                    else:
                        print(f"    -> X 下载失败 (可能是付费墙或非PDF): Type={content_type}")

                except Exception as e:
                    print(f"    -> X 请求出错: {e}")
            else:
                print(f"[{i + 1}] 跳过: 该条目没有 DOI")

            time.sleep(1)  # 礼貌延时

    except Exception as e:
        print(f"主程序出错: {e}")


if __name__ == '__main__':
    download_by_doi_construction()