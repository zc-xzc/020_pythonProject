import requests
import os
import time

# ================= 核心配置 =================
# 1. 使用您那个“能跑通”的 Open Access Key
API_KEY = '68dcdb2552f4b6b447f886735892aa22'

# 2. 接口地址 (Open Access)
BASE_URL = 'http://api.springernature.com/openaccess/json'

# 3. 您的搜索词
QUERY = 'Environmental Regulation'
SAVE_DIR = 'papers_final_download'  # 下载到这个新文件夹
MAX_PAPERS = 2  # 先下载5篇试试


# ===========================================

def download_papers_force_mode():
    # 创建文件夹
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)
        print(f"📂 创建下载目录: {SAVE_DIR}")

    print(f"🚀 开始搜索: '{QUERY}'...")
    print(f"🔑 使用 Key: {API_KEY[:5]}*** (Open Access)")

    # 1. 请求 API 获取文章列表
    params = {
        'q': QUERY,
        'api_key': API_KEY,
        'p': MAX_PAPERS,
        's': 1
    }

    try:
        response = requests.get(BASE_URL, params=params)
        response.raise_for_status()  # 如果 Key 不行，这里会报错
        data = response.json()

        records = data.get('records', [])
        print(f"✅ API 连接成功！获取到 {len(records)} 条文献信息。\n")

        # 2. 遍历每一篇文章
        for i, item in enumerate(records):
            title = item.get('title', 'untitled').replace('/', '-').replace(':', '-')
            doi = item.get('doi')  # 比如 "10.1007/s00467-025-06797-z"

            if not doi:
                print(f"[{i + 1}] 跳过: 没有 DOI 信息")
                continue

            # === 💡 核心黑科技：直接拼接 PDF 地址 ===
            # 无论 API 返回什么 URL，我们都用这个标准公式构造下载链接
            pdf_url = f"https://link.springer.com/content/pdf/{doi}.pdf"

            print(f"[{i + 1}] 正在下载: {title[:20]}...")
            print(f"    -> 构造地址: {pdf_url}")

            try:
                # 伪装成浏览器（防止被反爬虫拦截）
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                }

                # 请求下载
                file_response = requests.get(pdf_url, headers=headers, stream=True)

                # 检查是不是 PDF
                content_type = file_response.headers.get('Content-Type', '')
                if 'application/pdf' in content_type:
                    file_path = os.path.join(SAVE_DIR, f"{title}.pdf")
                    with open(file_path, 'wb') as f:
                        for chunk in file_response.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                    print(f"    -> 🎉 下载成功！")
                else:
                    # 如果返回的是 HTML，说明这篇不是免费的，或者是跳转页面
                    print(f"    -> ⚠️ 下载失败: 这是一个网页而不是PDF (可能是付费文章)")

            except Exception as e:
                print(f"    -> ❌ 网络错误: {e}")

            # 休息1秒，避免请求太快
            time.sleep(1)

    except requests.exceptions.HTTPError as err:
        if response.status_code == 401:
            print("❌ 严重错误: Key 依然报错 401。请确认您使用的是 Open Access Key。")
        else:
            print(f"❌ API 请求出错: {err}")
    except Exception as e:
        print(f"❌ 发生未知错误: {e}")

    print(f"\n🏁 任务完成。请查看文件夹: {SAVE_DIR}")


if __name__ == '__main__':
    download_papers_force_mode()