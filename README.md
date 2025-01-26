# 工廠智慧製造整合系統

一個結合 LINE Bot 的工廠智慧製造整合系統，實現 MES 流程自動化、物料管理、排程優化等功能。透過 Flask 框架建立後端服務，並使用 ngrok 進行穿透，實現完整的線上服務。

## 🌟 主要功能

### 1. LINE Bot 整合系統
- 前後端完整整合
- 即時訊息通知
- 使用者友善介面
- 多功能指令整合
- Flask 後端服務支援
- ngrok 穿透服務

### 2. MES 流程自動化
- 使用 Selenium 實現網頁自動化操作
- 自動登入、導航與表單填寫
- 批次處理大量工單
- 自動化數據擷取與回報
- 支援多瀏覽器環境 (Chrome, Firefox)
- 錯誤處理與自動重試機制
- Pandas 數據處理整合
- 自動化腳本執行

### 3. QR Code 管理系統
- 物料流水號 QR Code 產生器
- 一般用途 QR Code 請求產生
- 批次處理功能
- 自動編號系統

### 4. 智慧排程系統
- 最佳化自動排班表
- 人力資源配置優化
- 彈性調度功能
- 效率提升方案

## 💻 技術架構

- Python
- Flask Web 框架
- ngrok 穿透服務
- LINE Bot SDK
- Selenium WebDriver
- Chrome/Firefox WebDriver
- Pandas
- QR Code 生成工具
- RESTful API

## 🚀 安裝說明

1. 克隆專案
```bash
git clone [您的專案 URL]
```

2. 安裝相依套件
```bash
pip install -r requirements.txt
```

3. 安裝 WebDriver
```bash
# Chrome WebDriver
wget https://chromedriver.chromium.org/downloads/version-selection
# 或
# Firefox WebDriver (geckodriver)
wget https://github.com/mozilla/geckodriver/releases
```

4. 環境設定
```bash
# 設定 LINE Bot 驗證資訊
LINE_CHANNEL_SECRET=your_channel_secret
LINE_CHANNEL_ACCESS_TOKEN=your_access_token

# Flask 設定
FLASK_APP=app.py
FLASK_ENV=development

# ngrok 設定
NGROK_AUTH_TOKEN=your_ngrok_token

# Selenium 設定
CHROME_DRIVER_PATH=path/to/chromedriver
FIREFOX_DRIVER_PATH=path/to/geckodriver

# MES 系統設定
MES_URL=your_mes_system_url
MES_USERNAME=your_username
MES_PASSWORD=your_password
```

## 📖 使用說明

### 啟動服務
1. 啟動 Flask 服務
```bash
python app.py
```

2. 啟動 ngrok 穿透
```bash
ngrok http 5000
```

3. 更新 LINE Bot Webhook URL
- 將 ngrok 提供的 URL 設定到 LINE Developer Console
- Webhook URL 格式：`https://[ngrok-url]/callback`

### LINE Bot 指令列表
- `/help` - 顯示指令說明
- `/schedule` - 查看排班表
- `/qr` - 產生 QR Code

### MES 流程自動化操作
1. 配置 WebDriver
```python
# 在 config.py 中設定
SELENIUM_CONFIG = {
    'browser': 'chrome',  # 或 'firefox'
    'driver_path': 'path/to/webdriver',
    'headless': False,    # 是否使用無頭模式
    'implicit_wait': 10   # 隱式等待時間
}
```

2. 執行自動化流程
```bash
python mes_automation.py --process [process_name] --batch [batch_size]
```

3. 自動化功能
- 自動登入 MES 系統
- 批次處理工單
- 自動填寫表單
- 數據擷取與匯出
- 異常處理與記錄
- 自動化報表生成

### 錯誤處理機制
- 自動重試機制
- 異常日誌記錄
- 錯誤通知
- 系統狀態監控

## 🔧 配置說明

```python
# config.py

# Selenium 配置
SELENIUM_CONFIG = {
    'browser': 'chrome',
    'driver_path': '/path/to/chromedriver',
    'headless': False,
    'implicit_wait': 10,
    'page_load_timeout': 30,
    'retry_count': 3
}

# MES 系統配置
MES_CONFIG = {
    'url': 'https://your-mes-system.com',
    'username': 'your_username',
    'password': 'your_password',
    'auto_retry': True,
    'retry_interval': 5
}
```

## 📈 系統架構圖

[Client] <-> [ngrok] <-> [Flask Server] <-> [LINE Bot API]
                            |
                            ├── [Selenium WebDriver]
                            |       └── [MES System]
                            ├── [QR Generator]
                            └── [Schedule Optimizer] 

## 🤝 貢獻指南

1. Fork 此專案
2. 建立您的功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟一個 Pull Request

## 📝 版本記錄

- v1.2.1
    - Flask 整合優化
    - ngrok 穿透服務整合
    - Selenium MES 自動化流程優化
    - 新增自動重試機制
    - 改進錯誤處理
    - QR Code 生成器更新
    - 排程系統優化

## 👥 開發團隊

- 開發者：CHIEN
- 個人獨立開發專案

## 📄 授權資訊

此專案採用公開授權 (Public License)，歡迎自由使用、修改與分享。

## 📞 聯絡方式

- Email：bebw0717@gmail.com

## 💡 備註

本系統持續優化中，如有任何建議或問題，歡迎通過 Email 聯繫。 
