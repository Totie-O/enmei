import requests
from PIL import Image
import os
import math
import pandas as pd
from io import BytesIO
import time




def download_and_compress_image_plus(url, name, output_dir="D:\\款色图片文件夹", max_retries=3, high=200, wide=200):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    lst = []

    for attempt in range(max_retries):
        try:
            # 创建输出目录（如果不存在）
            os.makedirs(output_dir, exist_ok=True)
            
            # 流式下载图片
            response = requests.get(url, headers=headers, stream=True, timeout=10)
            response.raise_for_status()
            
            # 直接从内存中处理图片，避免不必要的磁盘写入
            img_data = BytesIO()
            for chunk in response.iter_content(1024):
                img_data.write(chunk)
            img_data.seek(0)
            
            # 打开并处理图片
            img = Image.open(img_data)
            
            # 转换图片模式为RGB（如果原始是RGBA或P模式）
            if img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')
            
            # 调整图片大小，保持宽高比
            img.thumbnail((high, wide), Image.Resampling.LANCZOS)
            
            # 准备保存路径
            save_path = os.path.join(output_dir, f"{name}.jpg")
            
            # 保存处理后的图片，优化压缩参数
            img.save(
                save_path,
                "JPEG",
                quality=95,
                optimize=True,
                progressive=True
            )
            
            print(f"图片处理成功：{name}")
            return True
            
        except requests.exceptions.RequestException as e:
            print(f"下载失败（尝试 {attempt + 1}/{max_retries}）：{name} | 错误：{str(e)}")
            if attempt == max_retries - 1:
                lst.append(name)
                return lst
            time.sleep(1)  # 等待1秒后重试
            
        except Exception as e:
            lst.append(name)
            print(f"处理失败：{name} | 错误：{str(e)}")
            return lst
        

if __name__ == "__main__":
    df = pd.read_excel(r"D:\桌面\款色图片.xlsx", sheet_name='Sheet14')

    for n, row in df.iterrows():
        download_and_compress_image_plus(row['图片'], row['款色'])

