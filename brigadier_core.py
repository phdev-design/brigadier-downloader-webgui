#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import subprocess
import plistlib
import re
import tempfile
import shutil
import datetime
import platform
import requests
import json
from urllib import request as urllib_request
from xml.dom import minidom

# 常數定義
VERSION = '0.2.8-webui'
SUCATALOG_URL = 'https://swscan.apple.com/content/catalogs/others/index-11-10.15-10.14-10.13-10.12-10.11-10.10-10.9-mountainlion-lion-snowleopard-leopard.merged-1.sucatalog'

def getCommandOutput(cmd):
    """執行一個命令並返回其 stdout。"""
    try:
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        out, _ = p.communicate()
        return out
    except OSError as e:
        print(f"執行命令時出錯 '{' '.join(cmd)}': {e}")
        return None

def getMachineModel():
    """返回本機的型號識別碼。"""
    try:
        if platform.system() == 'Windows':
            rawxml = getCommandOutput(['wmic', 'computersystem', 'get', 'model', '/format:RAWXML'])
            if not rawxml: return None
            dom = minidom.parseString(rawxml)
            nodes = dom.getElementsByTagName("VALUE")
            if nodes and nodes[0].childNodes:
                return nodes[0].childNodes[0].data
        elif platform.system() == 'Darwin':
            plistxml = getCommandOutput(['system_profiler', 'SPHardwareDataType', '-xml'])
            if not plistxml: return None
            plist = plistlib.loads(plistxml.encode('utf-8'))
            return plist[0]['_items'][0]['machine_model']
    except Exception as e:
        print(f"偵測型號時出錯: {e}")
    return None

def downloadFile(url, filename):
    """下載檔案，並使用 yield 回傳進度。"""
    def format_event(data):
        return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"

    yield format_event({'type': 'log', 'message': f"開始下載 {os.path.basename(filename)}..."})
    
    try:
        with requests.get(url, stream=True) as r:
            r.raise_for_status()
            total_size = int(r.headers.get('content-length', 0))
            bytes_downloaded = 0
            with open(filename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
                    bytes_downloaded += len(chunk)
                    if total_size > 0:
                        percent = (bytes_downloaded / total_size) * 100
                        yield format_event({
                            'type': 'progress',
                            'percent': round(percent, 2),
                            'downloaded': bytes_downloaded,
                            'total': total_size
                        })
        yield format_event({'type': 'log', 'message': "下載完成。"})
    except requests.exceptions.RequestException as e:
        yield format_event({'type': 'error', 'message': f"下載失敗: {e}"})
        raise

def sevenzipExtract(arcfile, command='e', out_dir=None):
    """使用 7-Zip 解壓縮存檔。"""
    sevenzip_binary = os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), "7-Zip", "7z.exe")
    if not os.path.exists(sevenzip_binary):
        payload = {'type': 'error', 'message': f'7-Zip 未在 {sevenzip_binary} 找到。請先安裝。'}
        yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"
        return

    cmd = [sevenzip_binary, command]
    if not out_dir:
        out_dir = os.path.dirname(arcfile)
    cmd.extend(["-o" + out_dir, "-y", arcfile])
    
    cmd_str = " ".join(cmd)
    payload = {'type': 'log', 'message': f'執行 7-Zip 命令: {cmd_str}'}
    yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"
    
    retcode = subprocess.call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if retcode:
        payload = {'type': 'error', 'message': f'7-Zip 命令失敗，返回碼: {retcode}'}
        yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"
    else:
        payload = {'type': 'log', 'message': '7-Zip 解壓縮成功。'}
        yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"


def run_brigadier(model, output_dir, product_id=None):
    """brigadier 的核心執行邏輯。"""
    def format_event(data):
        return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"

    yield format_event({'type': 'log', 'message': f"使用的 Mac 型號: {model}"})

    try:
        with urllib_request.urlopen(SUCATALOG_URL) as urlfd:
            data = urlfd.read()
    except urllib_request.URLError as e:
        yield format_event({'type': 'error', 'message': f"無法獲取軟體更新目錄: {e}"})
        return

    yield format_event({'type': 'log', 'message': "成功獲取軟體更新目錄。"})
    
    p = plistlib.loads(data)
    allprods = p.get('Products', {})
    bc_prods = [(pid, pdata) for pid, pdata in allprods.items() if 'BootCamp' in pdata.get('ServerMetadataURL', '')]

    pkg_data_list = []
    yield format_event({'type': 'log', 'message': "正在搜尋相符的 Boot Camp 產品..."})
    
    for prod_id_item, prod_data in bc_prods:
        dist_url = prod_data.get('Distributions', {}).get('English')
        if not dist_url: continue
        
        try:
            with urllib_request.urlopen(dist_url) as distfd:
                dist_data = distfd.read().decode('utf-8', errors='ignore')
            
            if re.search(model, dist_data, re.IGNORECASE):
                pkg_data_list.append({prod_id_item: prod_data})
                supported_models = sorted(list(set(re.findall(r"([a-zA-Z]{4,12}\d{1,2},\d{1,2})", dist_data))))
                yield format_event({
                    'type': 'product_found', 
                    'product_id': prod_id_item,
                    'post_date': prod_data.get('PostDate').strftime('%Y-%m-%d'),
                    'supported_models': ", ".join(supported_models)
                })
        except urllib_request.URLError:
            continue

    if not pkg_data_list:
        yield format_event({'type': 'error', 'message': f"找不到適用於型號 {model} 的 Boot Camp ESD。"})
        return

    pkg_data = None
    if product_id:
        yield format_event({'type': 'log', 'message': f"正在嘗試使用指定的產品 ID: {product_id}"})
        pkg_data = next((p_dict for p_dict in pkg_data_list if list(p_dict.keys())[0] == product_id), None)
        if not pkg_data:
            yield format_event({'type': 'error', 'message': f"指定的產品 ID {product_id} 無效或不適用於此型號。"})
            return
    elif len(pkg_data_list) == 1:
        pkg_data = pkg_data_list[0]
        yield format_event({'type': 'log', 'message': "找到唯一的產品，自動選擇。"})
    else:
        pkg_data = max(pkg_data_list, key=lambda p: list(p.values())[0].get('PostDate'))
        yield format_event({'type': 'log', 'message': f"找到多個產品，自動選擇最新的: {list(pkg_data.keys())[0]}"})

    if not pkg_data:
        yield format_event({'type': 'error', 'message': "無法選擇產品包。"})
        return

    pkg_id = list(pkg_data.keys())[0]
    pkg_url = pkg_data[pkg_id]['Packages'][0]['URL']
    yield format_event({'type': 'log', 'message': f"已選擇產品 ID: {pkg_id}"})

    landing_dir = os.path.join(output_dir, f'BootCampESD-{pkg_id}')
    if os.path.exists(landing_dir):
        yield format_event({'type': 'log', 'message': f"輸出路徑 {landing_dir} 已存在，正在移除..."})
        try:
            shutil.rmtree(landing_dir)
        except OSError as e:
            yield format_event({'type': 'error', 'message': f"移除舊目錄失敗: {e}"})
            return

    os.makedirs(landing_dir)
    yield format_event({'type': 'log', 'message': f"已建立目錄 {landing_dir}"})

    with tempfile.TemporaryDirectory(prefix="bootcamp-unpack_") as arc_workdir:
        pkg_dl_path = os.path.join(arc_workdir, os.path.basename(pkg_url))
        
        # 修正：確保將 pkg_url (下載網址) 傳遞給 downloadFile，而不是 pkg_dl_path (本地路徑)
        yield from downloadFile(pkg_url, pkg_dl_path)

        system = platform.system()
        if system == 'Windows':
            yield from sevenzipExtract(pkg_dl_path, command='x', out_dir=landing_dir)
        elif system == 'Darwin':
            yield format_event({'type': 'log', 'message': "正在展開 flat package..."})
            subprocess.call(['/usr/sbin/pkgutil', '--expand', pkg_dl_path, os.path.join(arc_workdir, 'pkg')])
            payload_path = os.path.join(arc_workdir, 'pkg', 'Payload')
            if os.path.exists(payload_path):
                yield format_event({'type': 'log', 'message': "正在解壓縮 Payload..."})
                subprocess.call(['/usr/bin/tar', '-xz', '-C', landing_dir, '-f', payload_path])
                yield format_event({'type': 'log', 'message': "Payload 解壓縮完成。"})
            else:
                yield format_event({'type': 'error', 'message': "在 package 中找不到 Payload。"})
        else: # Linux
            yield format_event({'type': 'log', 'message': "在 Linux 上嘗試使用 7z 解壓縮..."})
            subprocess.call(['7z', 'x', pkg_dl_path, f'-o{landing_dir}'])
    
    yield format_event({'type': 'log', 'message': f"解壓縮完成。儲存至 {landing_dir}"})

    yield format_event({'type': 'log', 'message': f"正在壓縮檔案至 ZIP..."})
    try:
        zip_filename_base = os.path.join(output_dir, f'BootCampESD-{pkg_id}')
        zip_path = shutil.make_archive(zip_filename_base, 'zip', landing_dir)
        zip_filename = os.path.basename(zip_path)
        yield format_event({'type': 'log', 'message': f"壓縮完成: {zip_filename}"})
        yield format_event({'type': 'zip_ready', 'filename': zip_filename})
    except Exception as e:
        yield format_event({'type': 'error', 'message': f"壓縮檔案時出錯: {e}"})
        return

    yield format_event({'type': 'done'})


def cleanup_files(output_dir, zip_filename):
    """刪除指定的 zip 檔案及其對應的解壓縮資料夾。"""
    try:
        zip_path = os.path.join(output_dir, zip_filename)
        folder_name = zip_filename.replace('.zip', '')
        folder_path = os.path.join(output_dir, folder_name)

        if os.path.exists(zip_path):
            os.remove(zip_path)
            print(f"已刪除 zip 檔案: {zip_path}")

        if os.path.exists(folder_path):
            shutil.rmtree(folder_path)
            print(f"已刪除資料夾: {folder_path}")
            
        return True
    except Exception as e:
        print(f"清理檔案時出錯: {e}")
        return False
