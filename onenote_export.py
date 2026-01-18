#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OneNote to Markdown Exporter

此脚本使用 Microsoft Graph API 将 OneNote 笔记本导出为 Markdown 格式，
并自动下载图片和附件。

Usage:
    python onenote_export.py

Author: gjican
License: MIT
"""

import os
import sys
import json
import re
import time
import requests
import msal
from markdownify import markdownify as md
from bs4 import BeautifulSoup

# 配置部分
# -----------------------------------------------------------------------------
# Microsoft Graph Client ID
# 这个 ID 是公开的，支持 Device Code Flow，适用于个人 Microsoft 账户。
# 如果你需要用于组织/学校账户，可能需要注册自己的 Azure App。
CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e' 

# 认证端点
# 对于个人账户使用 'consumers'，对于组织账户通常使用 'organizations' 或 'common'
AUTHORITY = 'https://login.microsoftonline.com/consumers'

# 请求的权限范围
SCOPES = ['Notes.Read', 'Notes.Read.All', 'User.Read']

# 导出目录名称
EXPORT_DIR = "OneNote_Export"
# -----------------------------------------------------------------------------

def get_access_token():
    app = msal.PublicClientApplication(
        CLIENT_ID, 
        authority=AUTHORITY
    )
    
    # 使用设备代码流 (Device Code Flow)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if 'user_code' not in flow:
        raise ValueError("无法初始化设备代码流: " + json.dumps(flow))
        
    print(f"\n>>> 请打开浏览器访问: {flow['verification_uri']}")
    print(f">>> 输入此代码: {flow['user_code']}")
    print(">>> 等待登录...\n")
    
    # 增加超时时间和重试
    max_retries = 3
    for i in range(max_retries):
        try:
            result = app.acquire_token_by_device_flow(flow)
            if "access_token" in result:
                return result['access_token']
            else:
                if "authorization_pending" in str(result):
                    continue
                print(f"错误: {result.get('error')}")
                print(f"描述: {result.get('error_description')}")
                sys.exit(1)
        except Exception as e:
            print(f"[登录连接错误] {str(e)} - 正在重试 ({i+1}/{max_retries})...")
            time.sleep(2)
            
    print("登录超时或网络中断，请检查网络后重试。")
    sys.exit(1)

def sanitize_filename(name):
    # 替换非法字符
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

def fetch_json(url, token, retries=5, use_pagination=False):
    headers = {'Authorization': 'Bearer ' + token}
    all_items = []
    
    # 如果开启分页，且 URL 里没有 top 参数，强制加上 top=20
    # 注意：只对获取列表的接口生效（notebooks, sections, pages）
    if use_pagination and "$top" not in url:
        separator = "&" if "?" in url else "?"
        url = f"{url}{separator}$top=20"

    current_url = url
    while current_url:
        data = None
        for i in range(retries):
            try:
                response = requests.get(current_url, headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    break
                elif response.status_code == 429: # Too Many Requests
                    wait_time = int(response.headers.get('Retry-After', 10))
                    print(f"      [429 限流] 等待 {wait_time} 秒...")
                    time.sleep(wait_time)
                    continue
                elif response.status_code >= 500: # Server Error
                    print(f"      [服务器错误 {response.status_code}] 重试 ({i+1}/{retries})...")
                    time.sleep(2 ** i) # 指数退避
                    continue
                else:
                    print(f"      [API 失败] {response.status_code} - {response.text[:100]}...")
                    return None
            except requests.exceptions.RequestException as e:
                print(f"      [网络错误] {str(e)}，重试中 ({i+1}/{retries})...")
                time.sleep(2 ** i)
                continue
        
        if not data:
            print("      [失败] 多次重试后无法获取数据")
            return None if not all_items else {'value': all_items}

        if 'value' in data:
            all_items.extend(data['value'])
            # 检查有没有下一页
            if '@odata.nextLink' in data:
                current_url = data['@odata.nextLink']
                print(f"      [分页] 获取下一页数据... (已获取 {len(all_items)} 条)")
            else:
                current_url = None
        else:
            # 如果不是列表结构（比如获取单个资源），直接返回
            return data

    return {'value': all_items}

def download_file(url, save_path, token, retries=3):
    headers = {'Authorization': 'Bearer ' + token}
    for i in range(retries):
        try:
            response = requests.get(url, headers=headers, stream=True)
            if response.status_code == 200:
                with open(save_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                return True
            elif response.status_code == 429:
                wait_time = int(response.headers.get('Retry-After', 10))
                time.sleep(wait_time)
                continue
            elif response.status_code >= 500:
                time.sleep(2 ** i)
                continue
        except Exception:
            time.sleep(2 ** i)
            continue
    return False

def process_page_content(page_id, token, assets_dir, retries=3):
    # 使用 includeInkML=true 获取墨迹信息（虽然主要还是靠 img 标签）
    url = f"https://graph.microsoft.com/v1.0/me/onenote/pages/{page_id}/content?includeIDs=true&includeInkML=true"
    headers = {'Authorization': 'Bearer ' + token}
    
    html_content = None
    for i in range(retries):
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                html_content = response.text
                break
            elif response.status_code == 429:
                wait_time = int(response.headers.get('Retry-After', 10))
                print(f"      [429 限流] 等待 {wait_time} 秒...")
                time.sleep(wait_time)
                continue
            elif response.status_code >= 500:
                time.sleep(2 ** i)
                continue
        except requests.exceptions.RequestException as e:
            print(f"      [网络错误] {str(e)}，重试中...")
            time.sleep(2 ** i)
            continue
            
    if not html_content:
        return None

    # 解析 HTML 处理图片和墨迹
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 创建 assets 目录
    if not os.path.exists(assets_dir):
        os.makedirs(assets_dir)

    # 查找所有图片 (img) 和对象 (object)
    # OneNote 的墨迹通常会以 <img data-src="..." /> 或 <object data="..." /> 的形式存在
    media_tags = soup.find_all(['img', 'object'])
    
    for idx, tag in enumerate(media_tags):
        # 获取下载链接
        # data-fullres-src 是高清图，src 是普通图
        src = tag.get('data-fullres-src') or tag.get('src') or tag.get('data')
        
        if not src or not src.startswith('http'):
            continue

        # 判断是否为附件 (Attachment)
        attachment_name = tag.get('data-attachment')
        is_attachment = bool(attachment_name)
            
        # 生成文件名
        if is_attachment:
            # 如果是附件，优先使用原文件名
            filename = sanitize_filename(attachment_name)
            # 防止文件名冲突，加个 ID 前缀
            filename = f"{page_id}_{filename}"
        else:
            # 图片/墨迹逻辑保持不变
            ext = '.png' 
            if 'image/jpeg' in str(tag): ext = '.jpg'
            elif 'application/pdf' in str(tag): ext = '.pdf' # 某些 PDF 打印件
            filename = f"{page_id}_asset_{idx}{ext}"

        save_path = os.path.join(assets_dir, filename)
        
        # 下载文件 (如果已存在且大小不为0，可以跳过下载，这里简单覆盖)
        if download_file(src, save_path, token):
            # 替换 HTML 中的链接为相对路径
            local_rel_path = f"assets/{filename}"
            
            if is_attachment:
                # 如果是附件，替换为一个 Markdown 链接： [文件名](路径)
                # 因为 markdownify 不会自动处理 object 为链接，我们需要手动把 object 换成 a 标签
                new_link = soup.new_tag("a", href=local_rel_path)
                new_link.string = f"📎 附件: {attachment_name}"
                tag.replace_with(new_link)
            elif tag.name == 'img':
                tag['src'] = local_rel_path
                # 移除 data-src 防止干扰
                if tag.has_attr('data-fullres-src'): del tag['data-fullres-src']
            elif tag.name == 'object':
                # 对于非附件的 object (可能是 PDF 打印件或墨迹)，转为 img
                new_img = soup.new_tag("img")
                new_img['src'] = local_rel_path
                tag.replace_with(new_img)
                
    return str(soup)

def main():
    if not os.path.exists(EXPORT_DIR):
        os.makedirs(EXPORT_DIR)
        
    print("正在获取访问令牌...")
    token = get_access_token()
    print("成功获取令牌！开始扫描笔记本...")
    
    # 1. 获取笔记本
    notebooks_url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
    notebooks_data = fetch_json(notebooks_url, token)
    
    if not notebooks_data or 'value' not in notebooks_data:
        print("未找到笔记本或权限不足。")
        return

    for nb in notebooks_data['value']:
        nb_name = sanitize_filename(nb['displayName'])
        print(f"\n处理笔记本: {nb_name}")
        nb_path = os.path.join(EXPORT_DIR, nb_name)
        if not os.path.exists(nb_path):
            os.makedirs(nb_path)
            
        # 2. 获取分区 (Sections)
        sections_url = f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{nb['id']}/sections"
        sections_data = fetch_json(sections_url, token)
        
        if not sections_data:
            continue
            
        all_sections = sections_data.get('value', [])

        for sec in all_sections:
            sec_name = sanitize_filename(sec['displayName'])
            print(f"  > 处理分区: {sec_name}")
            sec_path = os.path.join(nb_path, sec_name)
            assets_path = os.path.join(sec_path, "assets") # 每个分区一个 assets 文件夹
            
            if not os.path.exists(sec_path):
                os.makedirs(sec_path)
                
            # 3. 获取页面 (Pages)
            # $top=20 且只选择 id,title 字段，极大幅度降低 API 负载，避免 504
            pages_url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{sec['id']}/pages?$top=20&$select=id,title"
            pages_data = fetch_json(pages_url, token, use_pagination=True)
            
            if not pages_data or 'value' not in pages_data:
                continue
                
            all_pages = pages_data['value']
            
            if not all_pages:
                print("    [提示] 未发现笔记或获取失败")
                continue
                
            for page in all_pages:
                page_title = sanitize_filename(page['title'])
                if not page_title:
                    page_title = f"Untitled_{page['id']}"
                    
                # 检查文件是否已存在
                md_file_path = os.path.join(sec_path, f"{page_title}.md")
                assets_dir = os.path.join(sec_path, "assets")
                
                # 检查规则：
                # 1. 如果 MD 文件不存在，肯定要下载
                # 2. 如果 MD 文件存在，但内容里有 "http" 开头的图片链接（说明上次没下完图片），也要重新下载
                # 3. 如果 MD 文件存在且图片都是本地链接，则跳过
                should_download = True
                
                if os.path.exists(md_file_path):
                    try:
                        with open(md_file_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                            # 简单的判断：如果内容里没有 graph.microsoft.com 的图片链接，说明可能已经处理好了
                            # 或者更严格：检查 assets 目录里有没有对应图片
                            if "graph.microsoft.com" not in content and os.path.exists(assets_dir):
                                print(f"    [跳过] 已完整: {page_title}")
                                should_download = False
                            else:
                                print(f"    [补全] 发现未本地化图片: {page_title}")
                    except Exception:
                        pass # 读取错误则重新下载
                
                if not should_download:
                    continue

                print(f"    - 下载页面: {page_title}")
                
                # 下载处理后的 HTML 内容（包含图片下载逻辑）
                try:
                    processed_html = process_page_content(page['id'], token, assets_path)
                except Exception as e:
                    print(f"      [错误] 处理失败: {str(e)}")
                    continue
                    
                if processed_html:
                    # 转换为 Markdown
                    markdown_content = md(processed_html)
                    
                    # 保存文件
                    with open(md_file_path, 'w', encoding='utf-8') as f:
                        f.write(markdown_content)
                else:
                    print(f"      [失败] 无法下载内容")

    print(f"\n所有完成！笔记已保存在 {os.path.abspath(EXPORT_DIR)}")

if __name__ == '__main__':
    main()
