#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from flask import Flask, render_template, request, Response, stream_with_context, send_from_directory, jsonify, after_this_request
import brigadier_core

app = Flask(__name__, template_folder='templates')

# 定義下載目錄
DOWNLOADS_DIR = os.path.join(os.getcwd(), 'downloads')

@app.route('/')
def index():
    """渲染主頁面。"""
    default_model = brigadier_core.getMachineModel()
    return render_template('index.html', default_model=default_model)

@app.route('/run')
def run():
    """執行 brigadier 核心邏輯並串流輸出。"""
    model = request.args.get('model', '')
    product_id = request.args.get('product_id', '')
    
    if not model:
        def error_stream():
            # 使用 JSON 格式以保持一致性
            yield "data: {\"type\": \"error\", \"message\": \"錯誤：必須提供一個 Mac 型號。\"}\n\n"
        return Response(error_stream(), mimetype='text/event-stream')

    if not os.path.exists(DOWNLOADS_DIR):
        os.makedirs(DOWNLOADS_DIR)

    return Response(stream_with_context(brigadier_core.run_brigadier(model, DOWNLOADS_DIR, product_id)), mimetype='text/event-stream')

@app.route('/download/<path:filename>')
def download_zip(filename):
    """
    提供產生的 zip 檔案下載，並在下載完成後自動清理檔案。
    """
    # 使用 after_this_request 裝飾器來註冊一個請求結束後要執行的函式
    # 這可以確保只有在檔案成功發送後，清理工作才會開始
    @after_this_request
    def cleanup(response):
        brigadier_core.cleanup_files(DOWNLOADS_DIR, filename)
        return response

    # 從指定的目錄發送檔案
    return send_from_directory(DOWNLOADS_DIR, filename, as_attachment=True)

# 移除了舊的 /cleanup 路由，因為現在清理工作由 download_zip 函式處理

if __name__ == '__main__':
    # 將連接埠從 5000 更改為 5001
    app.run(host='0.0.0.0', port=5001, debug=True)
