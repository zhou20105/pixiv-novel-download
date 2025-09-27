import requests
import json
from docx import Document
import os
import re

def convert_cookies(input_file):
    """ 从文件中读取 cookies 并返回格式化后的字符串 """
    with open(input_file, 'r', encoding='utf-8') as f:
        cookie_string = f.read().strip()
    return cookie_string

def sanitize_filename(filename):
    """去除文件名中的非法字符"""
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def extract_novel_content(novel_id, cookies):
    """根据小说ID通过API获取小说标题和内容"""
    url = f"https://www.pixiv.net/ajax/novel/{novel_id}"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Cookie': cookies,
        'Referer': f'https://www.pixiv.net/novel/show.php?id={novel_id}',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        # 解析JSON响应
        data = response.json()
        
        # 检查响应是否成功
        if data.get('error'):
            print(f"API返回错误: {data.get('message', '未知错误')}")
            return None, None
        
        # 从body中提取数据
        body = data.get('body', {})
        
        # 提取标题
        novel_title = body.get('title', '')
        
        # 提取内容
        novel_content = body.get('content', '')
        
        # Unicode解码（如果需要）
        # JSON自动处理Unicode编码，所以通常不需要额外转换
        # 但如果内容中有转义的Unicode，这里会自动处理
        
        if not novel_title or not novel_content:
            print("未能获取到小说标题或内容")
            return None, None
            
        return novel_title, novel_content
        
    except requests.exceptions.RequestException as e:
        print(f"请求失败: {e}")
        return None, None
    except json.JSONDecodeError as e:
        print(f"JSON解析失败: {e}")
        return None, None
    except Exception as e:
        print(f"提取小说内容时发生错误: {e}")
        return None, None

def parse_novel_content(content):
    """解析小说内容，处理特殊标记"""
    # 处理换行标记
    content = content.replace('[newpage]', '\n\n')
    content = content.replace('[[rb:', '[')  # 处理ruby标记
    content = content.replace(']]', ']')
    
    # 分割段落
    paragraphs = content.split('\n')
    
    # 过滤空段落
    paragraphs = [p.strip() for p in paragraphs if p.strip()]
    
    return paragraphs

def save_novel_to_word(novel_title, content):
    """将提取的内容保存到Word文档"""
    doc = Document()
    doc.add_heading(novel_title, level=1)
    
    # 解析内容为段落
    paragraphs = parse_novel_content(content)
    
    for paragraph in paragraphs:
        if paragraph:
            doc.add_paragraph(paragraph)

    # 创建一个合法的文件名并保存文档
    sanitized_title = sanitize_filename(novel_title)
    file_name = f"{sanitized_title}.docx"
    save_path = os.path.join(os.getcwd(), file_name)
    doc.save(save_path)
    print(f"小说已保存到: {save_path}")

def main():
    # 从文件中读取cookies
    input_file = "cookie.txt"
    cookies = convert_cookies(input_file)
    
    print("Pixiv小说下载工具")
    print("=" * 50)
    
    while True:
        novel_id = input("\n请输入Pixiv小说ID（或输入 'exit' 退出）：").strip()
        
        if novel_id.lower() == 'exit':
            print("退出程序")
            break
        
        if not novel_id.isdigit():
            print("ID格式无效，请输入数字形式的小说ID")
            continue
        
        print(f"正在获取小说 ID: {novel_id}...")
        
        # 提取小说内容
        novel_title, novel_content = extract_novel_content(novel_id, cookies)
        
        if novel_title and novel_content:
            print(f"成功获取小说: {novel_title}")
            save_novel_to_word(novel_title, novel_content)
        else:
            print(f"无法获取小说 ID: {novel_id}")

if __name__ == "__main__":
    main()
