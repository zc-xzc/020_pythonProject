import requests
import os
import time

# ================= 关键修改区 =================
# 必须使用截图中的 "Meta API" Key，而不是 "开放获取API" Key
API_KEY = 'd5735d354312e9C8F6E96C10A0FF9178'  # <--- 已更正
# ============================================

QUERY = 'Environmental Regulation'  # 您的搜索词
SAVE_DIR = 'papers_pdf'  # 保存路径
MAX_PAPERS = 10  # 下载数量


def download_springer_pdfs():
    # 1. 创建保存文件夹
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)
        print(f"创建文件夹: {SAVE_DIR}")

    # 注意：这里使用的是 meta/v2/json 接口，必须搭配 Meta API Key
    base_url = 'http://api.springernature.com/meta/v2/json'
    print(f"开始搜索关键词: {QUERY}, 使用 Meta API Key...")

    downloaded_count = 0
    start_index = 1
    page_size = 50

    while downloaded_count < MAX_PAPERS:
        params = {
            'q': QUERY,
            'api_key': API_KEY,
            'p': page_size,
            's': start_index
        }

        try:
            # 发送请求
            print(f"正在请求 API (第 {start_index} 条开始)...")
            response = requests.get(base_url, params=params)

            # 检查是否成功
            if response.status_code == 401:
                print("错误：401 Unauthorized。请检查 Key 是否激活，或者是否选错了 Endpoint。")
                print(f"当前使用的 Key: {API_KEY}")
                return

            response.raise_for_status()
            data = response.json()

            if 'records' not in data or not data['records']:
                print("没有更多数据了，停止搜索。")
                break

            # 遍历并下载
            for item in data['records']:
                if downloaded_count >= MAX_PAPERS:
                    break

                title = item.get('title', 'untitled').replace('/', '-').replace(':', '-')

                # 寻找 PDF 链接
                pdf_url = None
                for url_item in item.get('url', []):
                    if url_item.get('format') == 'pdf':
                        pdf_url = url_item.get('value')
                        break

                if pdf_url:
                    print(f"[{downloaded_count + 1}] 正在下载: {title[:30]}...")
                    try:
                        # 尝试下载
                        pdf_response = requests.get(pdf_url, stream=True)

                        # 简单的内容检查
                        if 'application/pdf' not in pdf_response.headers.get('Content-Type', ''):
                            print(f"   -> 下载失败: 权限限制或非PDF文件")
                            continue

                        file_path = os.path.join(SAVE_DIR, f"{title}.pdf")
                        with open(file_path, 'wb') as f:
                            for chunk in pdf_response.iter_content(chunk_size=1024):
                                if chunk:
                                    f.write(chunk)
                        print(f"   -> 保存成功!")
                        downloaded_count += 1

                    except Exception as e:
                        print(f"   -> 下载出错: {e}")
                else:
                    print(f"   -> 跳过: 未找到PDF链接")

            start_index += page_size
            time.sleep(1)  # 避免请求过快

        except Exception as e:
            print(f"发生异常: {e}")
            break

    print(f"\n任务结束，共下载 {downloaded_count} 个文件。")


if __name__ == '__main__':
    download_springer_pdfs()