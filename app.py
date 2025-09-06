from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, make_response
import pandas as pd
import yfinance as yf
import os
from datetime import datetime, time
import re
import io
import logging
import json
import gspread
from google.oauth2.service_account import Credentials
import threading
import time

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev_secret_key")  # 使用環境變數

# 設置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 使用絕對路徑來存儲文件
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRANSACTION_FILE = os.path.join(BASE_DIR, "stock_transactions.csv")
STOCK_NAMES_FILE = os.path.join(BASE_DIR, "stock_names.csv")

# 全局緩存變量
TRANSACTIONS_CACHE = None
TRANSACTIONS_CACHE_TIME = None
TRANSACTIONS_CACHE_DURATION = 300  # 5分鐘緩存

# Google Sheets 設置
def setup_google_sheets():
    try:
        # 從環境變量獲取憑證（在 Render 上設置）
        creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
        if not creds_json:
            logger.error("未找到 Google Sheets 憑證環境變量")
            return None
        
        # 解析憑證
        scope = ['https://spreadsheets.google.com/feeds', 
                'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(
            json.loads(creds_json), scopes=scope)
        
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        logger.error(f"設置 Google Sheets 時出錯: {e}")
        return None

# 從 Google Sheets 讀取股票價格
def get_prices_from_google_sheet(client, sheet_name, worksheet_name="stock_names"):
    try:
        # 打開試算表
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        
        # 讀取所有數據
        data = sheet.get_all_records()
        
        # 轉換為字典格式：{股票代碼: 價格}
        prices = {}
        for row in data:
            if 'code' in row and 'price' in row:
                try:
                    # 確保代碼格式正確（帶有.TW或.TWO後綴）
                    code = str(row['code'])
                    # 如果代碼不包含後綴，嘗試添加.TW後綴
                    if not code.endswith(('.TW', '.TWO')):
                        code += '.TW'
                    
                    # 嘗試轉換為浮點數，如果失敗則跳過
                    price_value = float(row['price'])
                    prices[code] = price_value
                    logger.debug(f"從 Google Sheets 讀取股票價格: {code} = {price_value}")
                except (ValueError, TypeError) as e:
                    logger.warning(f"無法解析價格數據: {row}, 錯誤: {e}")
                    continue
        
        # 設置全局緩存
        if hasattr(fetch_stock_info, 'google_sheets_prices'):
            fetch_stock_info.google_sheets_prices = prices
        else:
            # 確保屬性存在
            fetch_stock_info.google_sheets_prices = prices
            
        logger.info(f"已從 Google Sheets 讀取 {len(prices)} 個股票價格")
        return prices
    except Exception as e:
        logger.error(f"從 Google Sheets 讀取數據時出錯: {e}")
        return {}

# 從 Google Sheets 讀取交易數據
def get_transactions_from_google_sheet(client, sheet_name, worksheet_name="交易紀錄"):
    try:
        # 打開試算表
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        
        # 讀取所有數據
        data = sheet.get_all_records()
        
        # 轉換為與原來相同的格式
        transactions = []
        for row in data:
            transactions.append({
                "Date": row.get("Date", ""),
                "Stock_Code": row.get("Stock_Code", ""),
                "Stock_Name": row.get("Stock_Name", ""),
                "Type": row.get("Type", ""),
                "Quantity": float(row.get("Quantity", 0)),
                "Price": float(row.get("Price", 0)),
                "Fee": float(row.get("Fee", 0)),
                "Tax": float(row.get("Tax", 0))
            })
        
        return transactions
    except Exception as e:
        logger.error(f"從 Google Sheets 讀取交易數據時出錯: {e}")
        return []

# 添加交易到 Google Sheets
def add_transaction_to_google_sheet(client, sheet_name, worksheet_name, transaction):
    try:
        # 打開試算表
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        
        # 直接添加新交易
        sheet.append_row([
            transaction["Date"],
            transaction["Stock_Code"],
            transaction["Stock_Name"],
            transaction["Type"],
            transaction["Quantity"],
            transaction["Price"],
            transaction["Fee"],
            transaction["Tax"]
        ])
        
        logger.info(f"已添加交易: {transaction['Stock_Code']} {transaction['Type']} {transaction['Quantity']}股")
        
        # 如果是買入交易，檢查並添加股票到 stock_names 工作表
        if transaction["Type"] == "Buy":
            # 提取股票代碼
            stock_code = transaction["Stock_Code"]
            
            # 檢查股票是否已存在於 stock_names 工作表
            if not check_stock_exists_in_names(client, sheet_name, stock_code):
                # 添加新股票到 stock_names 工作表
                add_stock_to_names_sheet(client, sheet_name, stock_code, transaction["Stock_Name"])
                logger.info(f"已將新股票 {stock_code} 添加到 stock_names 工作表")
            else:
                logger.info(f"股票 {stock_code} 已存在於 stock_names 工作表")
        
        return True
    except Exception as e:
        logger.error(f"添加交易到 Google Sheets 時出錯: {e}")
        return False

# 從 Google Sheets 刪除交易
def delete_transaction_from_google_sheet(client, sheet_name, worksheet_name, transaction_index):
    try:
        # 打開試算表
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        
        # 刪除指定行（行號從1開始，標題行是第一行，所以交易數據從第2行開始）
        # transaction_index 是交易列表中的索引，需要轉換為Google Sheets中的行號
        row_number = transaction_index + 2  # +2 是因為標題行(1)和0-based索引
        
        # 刪除行
        sheet.delete_rows(row_number)
        
        logger.info(f"已刪除交易，行號: {row_number}")
        return True
    except Exception as e:
        logger.error(f"從 Google Sheets 刪除交易時出錯: {e}")
        return False

# 檢查股票是否存在於 stock_names 工作表
def check_stock_exists_in_names(client, sheet_name, full_code):
    try:
        stock_names_sheet = client.open(sheet_name).worksheet("stock_names")
        records = stock_names_sheet.get_all_values()
        
        # 检查所有记录，查找匹配的代码
        for row in records:
            if len(row) > 0 and row[0] == full_code:  # 第一列是代码
                return True
        return False
    except gspread.exceptions.WorksheetNotFound:
        logger.warning("stock_names 工作表不存在")
        return False
    except Exception as e:
        logger.error(f"检查股票是否存在时出错: {e}")
        return False

# 添加新股票到 stock_names 工作表
def add_stock_to_names_sheet(client, sheet_name, full_code, name):
    try:
        # 尝试获取 stock_names 工作表，如果不存在则创建
        try:
            stock_names_sheet = client.open(sheet_name).worksheet("stock_names")
        except gspread.exceptions.WorksheetNotFound:
            # 创建更大的工作表（1000 行，10 列）
            stock_names_sheet = client.open(sheet_name).add_worksheet(title="stock_names", rows=1000, cols=10)
            # 添加标题行，注意顺序：code, price, name, pricenow
            stock_names_sheet.append_row(["code", "price", "name", "pricenow"])
            logger.info("已创建 stock_names 工作表")
        
        # 获取所有记录
        records = stock_names_sheet.get_all_values()
        
        # 找到第一个空行
        next_row = len(records) + 1
        
        # 检查是否超出网格限制
        if next_row > stock_names_sheet.row_count:
            # 增加行数
            stock_names_sheet.add_rows(100)
            logger.info(f"已增加 stock_names 工作表行数，当前行数: {stock_names_sheet.row_count}")
        
        # 构建公式
        if full_code.endswith('.TWO'):
            # 上櫃股票
            yahoo_code = full_code.replace(".TWO", ".TWO")
            formula = f'=IMPORTXML("https://tw.stock.yahoo.com/quote/{yahoo_code}","//*[@id=\'main-0-QuoteHeader-Proxy\']/div/div[2]/div[1]/div/span[1]")'
        else:
            # 上市股票
            yahoo_code = full_code.replace(".TW", "") + ".TW"
            formula = f'=IMPORTXML("https://tw.stock.yahoo.com/quote/{yahoo_code}","//*[@id=\'main-0-QuoteHeader-Proxy\']/div/div[2]/div[1]/div/span[1]")'
        
        # 使用批量更新
        batch_data = [
            {
                'range': f'A{next_row}',
                'values': [[full_code]]
            },
            {
                'range': f'B{next_row}',
                'values': [[f'=D{next_row}']]
            },
            {
                'range': f'C{next_row}',
                'values': [[name]]
            },
            {
                'range': f'D{next_row}',
                'values': [[formula]]
            }
        ]
        
        # 执行批量更新，使用 USER_ENTERED 选项
        stock_names_sheet.batch_update(batch_data, value_input_option='USER_ENTERED')
        
        logger.info(f"已将股票 {full_code} {name} 添加到 stock_names 工作表，行号: {next_row}")
        return True
    except Exception as e:
        logger.error(f"添加股票到 stock_names 工作表时出错: {e}")
        return False

# 在應用啟動時初始化 Google Sheets 連接
def initialize_google_sheets():
    try:
        client = setup_google_sheets()
        if client:
            # 從環境變量獲取試算表名稱
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
            
            # 檢查交易紀錄工作表是否存在，如果不存在則創建
            try:
                sheet = client.open(sheet_name).worksheet("交易紀錄")
            except gspread.exceptions.WorksheetNotFound:
                # 創建交易紀錄工作表
                sheet = client.open(sheet_name).add_worksheet(title="交易紀錄", rows=1000, cols=20)
                # 添加標題行
                sheet.append_row(["Date", "Stock_Code", "Stock_Name", "Type", "Quantity", "Price", "Fee", "Tax"])
            
            # 檢查 stock_names 工作表是否存在，如果不存在則創建
            try:
                stock_names_sheet = client.open(sheet_name).worksheet("stock_names")
            except gspread.exceptions.WorksheetNotFound:
                # 創建 stock_names 工作表
                stock_names_sheet = client.open(sheet_name).add_worksheet(title="stock_names", rows=1000, cols=10)
                # 添加標題行
                stock_names_sheet.append_row(["code", "name", "price", "pricenow"])
            
            # 新增：讀取價格數據
            get_prices_from_google_sheet(client, sheet_name, "stock_names")
            
            return True
        return False
    except Exception as e:
        logger.error(f"初始化 Google Sheets 時出錯: {e}")
        return False

# 定期更新 Google Sheets 數據
def schedule_google_sheets_update(interval_minutes=30):
    def update():
        while True:
            try:
                time.sleep(interval_minutes * 60)
                client = setup_google_sheets()
                if client:
                    sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
                    get_prices_from_google_sheet(client, sheet_name, "stock_names")
                    logger.info("已更新 Google Sheets 價格數據")
            except Exception as e:
                logger.error(f"定期更新 Google Sheets 數據時出錯: {e}")
    
    # 啟動後台線程
    thread = threading.Thread(target=update)
    thread.daemon = True
    thread.start()

# 獲取交易數據（使用緩存）
def get_transactions():
    global TRANSACTIONS_CACHE, TRANSACTIONS_CACHE_TIME
    
    current_time = time.time()
    if (TRANSACTIONS_CACHE is not None and 
        TRANSACTIONS_CACHE_TIME is not None and
        current_time - TRANSACTIONS_CACHE_TIME < TRANSACTIONS_CACHE_DURATION):
        logger.info("使用緩存的交易數據")
        return TRANSACTIONS_CACHE
    
    try:
        client = setup_google_sheets()
        if client:
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
            transactions = get_transactions_from_google_sheet(client, sheet_name, "交易紀錄")
            TRANSACTIONS_CACHE = transactions
            TRANSACTIONS_CACHE_TIME = current_time
            logger.info(f"從 Google Sheets 讀取 {len(transactions)} 筆交易數據")
            return transactions
        else:
            logger.warning("無法連接到 Google Sheets，返回空交易列表")
            return []
    except Exception as e:
        logger.error(f"獲取交易數據時出錯: {e}")
        return []

# Load stock names from CSV with encoding fallback
def load_stock_names():
    try:
        if not os.path.exists(STOCK_NAMES_FILE):
            logger.warning(f"{STOCK_NAMES_FILE} 不存在，使用空映射")
            return {}
        try:
            df = pd.read_csv(STOCK_NAMES_FILE, encoding='utf-8-sig')
        except UnicodeDecodeError:
            logger.warning("無法以 utf-8-sig 編碼讀取 stock_names.csv，嘗試 big5")
            try:
                df = pd.read_csv(STOCK_NAMES_FILE, encoding='big5')
            except UnicodeDecodeError:
                logger.error("無法以 utf-8-sig 或 big5 編碼讀取 stock_names.csv，請檢查檔案編碼")
                return {}
        expected_columns = ["Code", "Name", "Market"]
        if list(df.columns) != expected_columns:
            logger.error(f"{STOCK_NAMES_FILE} 格式錯誤，應包含欄位: {expected_columns}")
            return {}
        stock_names = {}
        for _, row in df.iterrows():
            try:
                code = str(row["Code"])
                market = row["Market"]
                name = row["Name"]
                stock_names[(code, market)] = name
                logger.debug(f"股票映射: 代碼={code}, 市場={market}, 名稱={name}")
            except Exception as e:
                logger.warning(f"跳過無效行: {row.to_dict()}, 錯誤: {e}")
        logger.info(f"成功載入 {len(stock_names)} 個股票名稱")
        return stock_names
    except Exception as e:
        logger.error(f"載入 {STOCK_NAMES_FILE} 失敗: {e}")
        return {}

# Fetch stock info - 使用 Google Sheets 數據
def fetch_stock_info(full_code):
    # 從完整代碼中提取基本信息
    if full_code.endswith(".TWO"):
        code = full_code.split('.')[0]
        is_otc = True
        market_key = "TWO"
    else:
        code = full_code.split('.')[0]
        is_otc = False
        market_key = "TWSE"
    
    # 確保 google_sheets_prices 屬性存在
    if not hasattr(fetch_stock_info, 'google_sheets_prices'):
        fetch_stock_info.google_sheets_prices = {}
    
    # 確保緩存存在
    if not hasattr(fetch_stock_info, 'cache'):
        fetch_stock_info.cache = {}
    
    # 使用緩存來減少 API 請求
    cache_key = full_code
    current_time = time.time()
    
    # 檢查緩存是否存在且未過期（30分鐘）
    cached_data = fetch_stock_info.cache.get(cache_key)
    if cached_data and current_time - cached_data['timestamp'] < 1800:  # 30分鐘緩存
        logger.info(f"使用緩存的股票數據: {cache_key}")
        return cached_data['data']
    
    # 從本地 CSV 獲取股票名稱
    stock_names = load_stock_names()
    name_key = (str(code), market_key)
    name = stock_names.get(name_key, "未知名稱")
    
    # 嘗試從 Google Sheets 獲取價格 - 使用完整代碼（帶後綴）
    price = 0
    if fetch_stock_info.google_sheets_prices:
        price = fetch_stock_info.google_sheets_prices.get(full_code, 0)
    
    # 如果 Google Sheets 沒有數據，嘗試使用 Yahoo Finance
    if price == 0:
        try:
            stock = yf.Ticker(full_code)
            hist = stock.history(period="1d")
            if not hist.empty:
                price = hist["Close"].iloc[-1]
                logger.info(f"從 Yahoo Finance 獲取 {full_code} 價格: {price}")
        except Exception as e:
            logger.error(f"從 Yahoo Finance 獲取 {full_code} 價格失敗: {e}")
            price = 0
    
    result = {"price": round(price, 2), "name": name}
    
    # 更新緩存
    fetch_stock_info.cache[cache_key] = {
        'timestamp': current_time,
        'data': result
    }
    
    return result

# Calculate portfolio summary
def get_portfolio_summary(transactions=None):
    if transactions is None:
        transactions = get_transactions()
        
    if not transactions:
        return [], 0, 0, 0, 0, 0

    # 使用字典來跟踪每個股票的狀態
    stock_status = {}
    
    for row in transactions:
        # 提取不帶後綴的股票代碼
        code_without_suffix = row["Stock_Code"].split('.')[0]
        full_code = row["Stock_Code"]
        
        if full_code not in stock_status:
            stock_status[full_code] = {
                "name": row["Stock_Name"],
                "quantity": 0,
                "total_buy_cost": 0,  # 總買入成本
                "total_sell_revenue": 0,  # 總賣出收入
                "total_fee_tax": 0,  # 總手續費和稅
                "buy_quantity": 0,  # 總買入股數
                "sell_quantity": 0,  # 總賣出股數
                "realized_profit": 0,  # 已實現損益
                "is_otc": row["Stock_Code"].endswith(".TWO")
            }
        
        if row["Type"] == "Buy":
            stock_status[full_code]["quantity"] += row["Quantity"]
            stock_status[full_code]["total_buy_cost"] += row["Quantity"] * row["Price"] + row["Fee"] + row["Tax"]
            stock_status[full_code]["buy_quantity"] += row["Quantity"]
            stock_status[full_code]["total_fee_tax"] += row["Fee"] + row["Tax"]
        else:  # Sell
            stock_status[full_code]["quantity"] -= row["Quantity"]
            stock_status[full_code]["total_sell_revenue"] += row["Quantity"] * row["Price"] - row["Fee"] - row["Tax"]
            stock_status[full_code]["sell_quantity"] += row["Quantity"]
            stock_status[full_code]["total_fee_tax"] += row["Fee"] + row["Tax"]
            
            # 計算已實現損益
            avg_buy_price = stock_status[full_code]["total_buy_cost"] / stock_status[full_code]["buy_quantity"] if stock_status[full_code]["buy_quantity"] > 0 else 0
            realized_profit = (row["Price"] - avg_buy_price) * row["Quantity"] - row["Fee"] - row["Tax"]
            stock_status[full_code]["realized_profit"] += realized_profit

    result = []
    total_cost = 0
    total_market_value = 0
    total_unrealized_profit = 0
    total_realized_profit = 0
    total_quantity = 0

    for full_code, data in stock_status.items():
        code_without_suffix = full_code.split('.')[0]
        
        # 獲取當前股價（只有當持有股數大於0時才需要）
        current_price = 0
        if data["quantity"] > 0:
            stock_info = fetch_stock_info(full_code)
            current_price = stock_info["price"]
        
        # 計算市值和未實現損益
        market_value = data["quantity"] * current_price
        avg_buy_price = data["total_buy_cost"] / data["buy_quantity"] if data["buy_quantity"] > 0 else 0
        unrealized_profit = (current_price - avg_buy_price) * data["quantity"] if data["quantity"] > 0 else 0
        
        # 計算總成本（只計算當前持有的部分）
        current_cost = avg_buy_price * data["quantity"] if data["quantity"] > 0 else 0
        
        # 累加總計
        total_quantity += data["quantity"]
        total_cost += current_cost
        total_market_value += market_value
        total_unrealized_profit += unrealized_profit
        total_realized_profit += data["realized_profit"]

        result.append({
            "Stock_Code": code_without_suffix,  # 不帶後綴的代碼，用於顯示
            "Stock_Name": data["name"],
            "Quantity": int(data["quantity"]),
            "Avg_Buy_Price": round(avg_buy_price, 2),
            "Current_Price": round(current_price, 2),
            "Total_Cost": int(current_cost),
            "Market_Value": int(market_value),
            "Unrealized_Profit": int(unrealized_profit),
            "Realized_Profit": int(data["realized_profit"]),
            "Full_Code": full_code  # 保存完整代碼供其他用途
        })

    return result, total_quantity, total_cost, total_market_value, total_unrealized_profit, total_realized_profit

# 主頁面路由
@app.route("/", methods=["GET", "POST"])
def index():
    global TRANSACTIONS_CACHE
    
    initialize_google_sheets()
    error = None
    stock_name = None
    default_date = datetime.now().strftime("%Y-%m-%d")
    add_transaction_message = None
    update_all_prices_message = None
    delete_transaction_message = None

    # 獲取交易數據和投資組合摘要
    transactions = get_transactions()
    summary, total_quantity, total_cost, total_market_value, total_unrealized_profit, total_realized_profit = get_portfolio_summary(transactions)

    if request.method == "POST":
        action = request.form.get("action")
        
        if action == "add_transaction":
            try:
                date = request.form.get("date", default_date)
                code = request.form.get("code", "").strip()
                name = request.form.get("name", "").strip() or "未知股票"
                market = request.form.get("market", "TWSE")
                trans_type = request.form.get("type", "Buy")
                quantity = request.form.get("quantity")
                price = request.form.get("price")
                
                # Validate inputs
                if not code:
                    error = "股票代碼不能為空"
                elif not quantity or float(quantity) <= 0:
                    error = "股數必須為正數"
                elif float(quantity) % 1000 != 0:
                    error = "股數必須為1000的倍數"
                elif not price or float(price) <= 0:
                    error = "每股價格必須為正數"
                else:
                    quantity = float(quantity)
                    price = float(price)
                    # 自動計算手續費和交易稅
                    fee = max(20, price * quantity * 0.001425)
                    tax = price * quantity * 0.003 if trans_type == "Sell" else 0

                    code_with_suffix = f"{code}.TWO" if market == "TWO" else f"{code}.TW"
                    new_transaction = {
                        "Date": date,
                        "Stock_Code": code_with_suffix,
                        "Stock_Name": name,
                        "Type": trans_type,
                        "Quantity": quantity,
                        "Price": price,
                        "Fee": fee,
                        "Tax": tax
                    }
                    
                    # 添加到 Google Sheets
                    client = setup_google_sheets()
                    if client:
                        sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
                        if add_transaction_to_google_sheet(client, sheet_name, "交易紀錄", new_transaction):
                            # 清除交易緩存
                            TRANSACTIONS_CACHE = None
                            add_transaction_message = "交易已新增！"
                            
                            # 重新獲取交易數據
                            transactions = get_transactions()
                            summary, total_quantity, total_cost, total_market_value, total_unrealized_profit, total_realized_profit = get_portfolio_summary(transactions)
                        else:
                            error = "無法將交易添加到 Google Sheets"
                    else:
                        error = "無法連接到 Google Sheets"
            except ValueError as e:
                error = f"輸入無效: {str(e)}。請確保股數和價格為有效數字"
            except Exception as e:
                error = f"新增交易時發生錯誤: {str(e)}"
                logger.error(f"新增交易失敗: {e}")
        
        elif action == "update_all_prices":
            try:
                client = setup_google_sheets()
                if client:
                    sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
                    # 强制重新从 Google Sheets 读取价格数据
                    get_prices_from_google_sheet(client, sheet_name, "stock_names")
                    
                    # 清除股票信息缓存，强制重新获取所有股票的最新价格
                    if hasattr(fetch_stock_info, 'cache'):
                        fetch_stock_info.cache = {}
                    
                    # 重新计算投资组合摘要
                    summary, total_quantity, total_cost, total_market_value, total_unrealized_profit, total_realized_profit = get_portfolio_summary(transactions)
                    
                    update_all_prices_message = "所有股價已更新！"
                    logger.info("已强制更新所有股价")
                else:
                    error = "無法連接到 Google Sheets"
            except Exception as e:
                error = f"更新股價時出錯: {str(e)}"
                logger.error(f"更新所有股價失敗: {e}")
        
        elif action == "delete_transaction":
            try:
                transaction_index = request.form.get("transaction_index")
                if transaction_index is not None:
                    transaction_index = int(transaction_index)
                    
                    client = setup_google_sheets()
                    if client:
                        sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
                        if delete_transaction_from_google_sheet(client, sheet_name, "交易紀錄", transaction_index):
                            # 清除交易緩存
                            TRANSACTIONS_CACHE = None
                            delete_transaction_message = "交易已刪除！"
                            
                            # 重新獲取交易數據
                            transactions = get_transactions()
                            summary, total_quantity, total_cost, total_market_value, total_unrealized_profit, total_realized_profit = get_portfolio_summary(transactions)
                        else:
                            error = "無法從 Google Sheets 刪除交易"
                    else:
                        error = "無法連接到 Google Sheets"
            except Exception as e:
                error = f"刪除交易時發生錯誤: {str(e)}"
                logger.error(f"刪除交易失敗: {e}")

    # 渲染模板（適用於 GET 和 POST 請求）
    return render_template(
        "index.html",
        transactions=transactions,
        summary=summary,
        total_quantity=total_quantity,
        total_cost=total_cost,
        total_market_value=total_market_value,
        total_unrealized_profit=total_unrealized_profit,
        total_realized_profit=total_realized_profit,
        error=error,
        stock_name=stock_name,
        default_date=default_date,
        add_transaction_message=add_transaction_message,
        update_all_prices_message=update_all_prices_message,
        delete_transaction_message=delete_transaction_message
    )

# 獲取股票名稱
@app.route("/fetch_stock_name", methods=["POST"])
def fetch_stock_name():
    code = request.form.get("code", "").strip()
    market = request.form.get("market", "TWSE")
    logger.info(f"收到查詢請求: 代碼={code}, 市場={market}")
    
    if not code:
        response = jsonify({"error": "請輸入股票代碼"})
        response.headers["Content-Type"] = "application/json; charset=utf-8"
        return response
    
    # 直接從本地 CSV 獲取股票名稱
    stock_names = load_stock_names()
    
    # 根據市場選擇正確的鍵
    if market == "TWO":
        market_key = "TWO"
    else:
        market_key = "TWSE"
    
    name_key = (str(code), market_key)
    logger.info(f"查找的鍵: {name_key}")
    
    # 記錄所有可用的鍵以便調試
    available_keys = list(stock_names.keys())
    logger.info(f"可用的鍵數量: {len(available_keys)}")
    if available_keys:
        logger.info(f"前幾個可用鍵: {available_keys[:5]}")
    
    name = stock_names.get(name_key, "")
    
    # 特殊处理：对于债券代码，可能需要不同的查找方式
    if not name and code.endswith('B') and market == "TWO":
        # 尝试去掉债券标识 'B' 再查找
        code_without_b = code.rstrip('B')
        name_key_without_b = (code_without_b, market_key)
        name = stock_names.get(name_key_without_b, "")
        if name:
            logger.info(f"找到去掉債券標識後的名稱: {name}")
    
    if not name:
        # 嘗試不區分市場查找
        for key, value in stock_names.items():
            if key[0] == str(code):
                name = value
                logger.info(f"找到不區分市場的名稱: {name}")
                break
        
        if not name:
            logger.error(f"無法找到股票 {code} 的名稱，查找的鍵: {name_key}")
            response = jsonify({"error": f"無法找到股票 {code} 的名稱，請手動輸入名稱"})
            response.headers["Content-Type"] = "application/json; charset=utf-8"
            return response
    
    logger.info(f"返回股票名稱: {name}")
    response = jsonify({"name": name, "is_english": not re.search(r'[\u4e00-\u9fff]', name)})
    response.headers["Content-Type"] = "application/json; charset=utf-8"
    return response

# 匯出交易紀錄
@app.route("/export_transactions")
def export_transactions():
    try:
        client = setup_google_sheets()
        if client:
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
            transactions = get_transactions_from_google_sheet(client, sheet_name, "交易紀錄")
            
            # 轉換為 DataFrame 並導出為 CSV
            df = pd.DataFrame(transactions)
            output = io.StringIO()
            df.to_csv(output, index=False, encoding='utf-8-sig')
            output.seek(0)
            
            return send_file(
                io.BytesIO(output.getvalue().encode("utf-8-sig")),
                mimetype="text/csv; charset=utf-8",
                as_attachment=True,
                download_name=f"exported_transactions_{datetime.now().strftime('%Y%m%d')}.csv"
            )
        else:
            flash("無法連接到 Google Sheets", "error")
            return redirect(url_for("index"))
    except Exception as e:
        flash(f"匯出失敗: {e}", "error")
        return redirect(url_for("index"))

# 初始化 Google Sheets 並啟動定期更新
initialize_google_sheets()
schedule_google_sheets_update(30)  # 每30分鐘更新一次

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
