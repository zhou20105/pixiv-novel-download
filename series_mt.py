import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import requests
import json
import os
import re
from concurrent.futures import ProcessPoolExecutor
from docx import Document

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

def save_novel_to_word(novel_title, content, series_folder, chapter_number):
    """将提取的内容保存到Word文档"""
    doc = Document()
    doc.add_heading(f"第 {chapter_number} 章: {novel_title}", level=1)
    
    # 解析内容为段落
    paragraphs = parse_novel_content(content)
    
    for paragraph in paragraphs:
        if paragraph:
            doc.add_paragraph(paragraph)

    # 创建一个合法的文件名并保存文档
    sanitized_title = sanitize_filename(novel_title)
    file_name = f"第{chapter_number}章_{sanitize_filename(novel_title)}.docx"
    save_path = os.path.join(series_folder, file_name)
    doc.save(save_path)
    print(f"小说已保存到: {save_path}")


# 提取系列章节URL
def extract_novel_ids(series_id, cookies):
    novel_ids = []
    series_url = f"https://www.pixiv.net/ajax/novel/series/{series_id}/content_titles"
    headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Cookie': cookies,
                'Referer': f'https://www.pixiv.net/novel/series/{series_id}',
                'Accept': 'application/json, text/plain, */*',
                'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            }
    try:
        response = requests.get(series_url, headers=headers)
        response.raise_for_status()
        # 解析JSON响应
        data = response.json()
    
        # 检查响应是否成功
        if data.get('error'):
            print(f"API返回错误: {data.get('message', '未知错误')}")
            return None, None
    
        # 从body中提取数据
        body = data.get('body', {})
    
        for item in body:
            novel_id = item.get('id')
            if novel_id:
                novel_ids.append(novel_id)
    except requests.exceptions.RequestException as e:
        print(f"请求失败: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"JSON解析失败: {e}")
        return None
    except Exception as e:
        print(f"提取章节URL时发生错误: {e}")
        return None
    return novel_ids


# 获取系列名称并创建文件夹
def create_series_folder(series_title, output_path=None):
    try:
        
        # 创建文件夹
        if output_path:
            folder_path = os.path.join(output_path, series_title)
        else:
            folder_path = series_title
        try:
            os.makedirs(folder_path, exist_ok=True)
            return folder_path
        except Exception as e:
            print(f"Error creating folder: {e}")
            return None
    except Exception as e:
        print(f"Error creating folder: {e}")
        return None


def process_chapter(novel_id, chapter_number, series_folder, cookies):
    """ 处理每个章节的下载 """
    try:
        novel_title, content = extract_novel_content( novel_id, cookies)
        if novel_title and content:
            save_novel_to_word(novel_title, content, series_folder, chapter_number)
    except Exception as e:
        print(f"Error processing chapter {chapter_number}: {e}")


def download_series(series_id, status_label, progress_bar, output_label, output_path):
    input_file = "cookie.txt"
    cookies = convert_cookies(input_file)
    
    status_label.config(text="Downloading...")
    
    try:
        # 
        series_url = f"https://www.pixiv.net/ajax/novel/series/{series_id}"
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Cookie': cookies,
            'Referer': f'https://www.pixiv.net/novel/series/{series_id}',
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        }
        
        try:
            response = requests.get(series_url, headers=headers)
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
            series_title = body.get('title', '')
            
        except requests.exceptions.RequestException as e:
            status_label.config(text=f"请求失败: {e}")
            return
        except json.JSONDecodeError as e:
            status_label.config(text=f"JSON解析失败: {e}")
            return
        except Exception as e:
            status_label.config(text=f"提取系列标题时发生错误: {e}")
            return
        
        series_folder = create_series_folder(series_title, output_path)
        if not series_folder:
            status_label.config(text="创建系列文件夹失败，跳过此系列。")
            return
        
        output_label.config(text=f"Output folder: {series_folder}")
        novel_ids = extract_novel_ids(series_id, cookies)
        
        total_chapters = len(novel_ids)
        progress_bar['maximum'] = total_chapters
        
        # 使用进程池并行下载章节，确保每个进程使用独立的浏览器实例
        with ProcessPoolExecutor(max_workers=5) as executor:
            futures = [
                executor.submit(process_chapter, novel_id, i + 1, series_folder, cookies)
                for i, novel_id in enumerate(novel_ids)
            ]
            # 等待所有任务完成
            for i, future in enumerate(futures):
                future.result()
                progress_bar['value'] = i + 1
                progress_bar.update()
        
        status_label.config(text="Completed")
    
    except Exception as e:
        status_label.config(text=f"发生错误：{e}")


def main():
    window = tk.Tk()
    window.title("Pixiv Series Downloader")
    
    series_id_label = ttk.Label(window, text="Pixiv Series ID:")
    series_id_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
    
    series_id_entry = ttk.Entry(window)
    series_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    output_path_label = ttk.Label(window, text="Output Path:")
    output_path_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")

    output_path_entry = ttk.Entry(window)
    output_path_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    
    # Button to open a folder selection dialog
    def select_folder():
        folder_selected = filedialog.askdirectory()
        if folder_selected:  # If a folder is selected, update the entry
            output_path_entry.delete(0, tk.END)
            output_path_entry.insert(0, folder_selected)
    
    browse_button = ttk.Button(window, text="Browse", command=select_folder)
    browse_button.grid(row=1, column=2, padx=5, pady=5, sticky="w")
    
    status_label = ttk.Label(window, text="Ready")
    status_label.grid(row=2, column=0, columnspan=3, padx=5, pady=5)
    
    progress_bar = ttk.Progressbar(window, orient="horizontal", length=300, mode="determinate")
    progress_bar.grid(row=3, column=0, columnspan=3, padx=5, pady=5)
    
    output_label = ttk.Label(window, text="")
    output_label.grid(row=4, column=0, columnspan=3, padx=5, pady=5)
    
    def start_download():
        series_id = series_id_entry.get().strip()
        output_path = output_path_entry.get().strip()
        if not series_id:
            messagebox.showerror("Error", "Please enter a Pixiv Series ID.")
            return
        # If output_path is not provided, set a default value
        if not output_path:
            output_path = os.getcwd()  # Default to the current working directory
            messagebox.showinfo("Info", f"No output path provided. Using default path: {output_path}")
        
        download_series(series_id, status_label, progress_bar, output_label, output_path)
    
    start_button = ttk.Button(window, text="Start Download", command=start_download)
    start_button.grid(row=5, column=0, columnspan=3, padx=5, pady=10)
    
    window.mainloop()


if __name__ == "__main__":
    main()
