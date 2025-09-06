#!/usr/bin/env bash
# 啟動應用的腳本
echo "啟動股票投資管理應用..."
# 設置台灣時區
export TZ=Asia/Taipei
# 設置更長的超時時間
gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 2 --timeout 120 --preload