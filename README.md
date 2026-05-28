# 📂 金卡檔案 OCR → Excel 轉換工具

本專案是一個基於 **Streamlit** 與 **Gemini API** 的 Web 工具，能將 PDF/影像檔（申請書、護照、學歷證明、工作經歷等）自動進行 OCR，並提取結構化資料，輸出為 **美化後的 Excel**，協助數位金卡審查流程自動化。

---

## ✨ 功能特色
- 🔍 **OCR 辨識**：使用 Gemini 2.5 Flash 模型，支援 PDF 及圖片檔案。
- 📊 **自動結構化**：依指定欄位抽取資料並格式化為 JSON。
- 📥 **匯出美化 Excel**：
  - 固定欄寬、標題藍底白字加粗
  - 長文字自動換行與列高統一
  - 字體統一 14pt
- 🌐 **Web 部署**：支援 **Streamlit Cloud** 一鍵部署，方便內部或跨部門使用。

---

## 🛠️ 技術架構
- **前端框架**：[Streamlit](https://streamlit.io/)
- **OCR 模型**：Google [Gemini 2.5 Flash](https://ai.google.dev/)
- **資料處理**：Pandas
- **Excel 美化**：OpenPyXL

---

## 📦 安裝與本地執行

### 1️⃣ 下載專案
```bash
git clone https://github.com/IanChen-adi/one_click_goldencard_review.git
cd one_click_goldencard_review
```

### 2️⃣ 安裝依賴套件
```bash
pip install -r requirements.txt
```
### 3️⃣ 設定環境變數（本地）
在專案根目錄建立 .env 檔案，內容如下：
```bash
GEMINI_API_KEY=你的_Gemini_API_Key
```
### 4️⃣ 執行 Streamlit App
```bash
streamlit run auto_aiV3.py
```
