# Marp 簡報轉換工具使用說明

這是一個將純文本文件轉換為精美簡報的工具，支援 Markdown、PDF 和 PPTX 格式輸出。本工具基於 Marp CLI 開發，提供了命令行和圖形界面兩種使用方式。

## 安裝要求

1. Python 3.6 或更高版本
2. Node.js 和 npm（用於安裝 Marp CLI）
3. Marp CLI：執行以下命令安裝
   ```bash
   npm install -g @marp-team/marp-cli
   ```

## 使用方法

### 圖形界面模式

1. 直接運行程式：
   ```bash
   python3 purepython-txt_to_marp_ppt.py
   ```

2. 操作步驟：
   - 點擊「瀏覽」選擇輸入的 TXT 文件
   - 選擇輸出目錄（默認為輸入文件所在目錄）
   - 選擇簡報主題（default、gaia 或 uncover）
   - 勾選需要的輸出格式（PDF 和/或 PPTX）
   - 點擊「開始轉換」

### 命令行模式

基本用法：
```bash
python3 purepython-txt_to_marp_ppt.py input.txt [選項]
```

可用選項：
- `-o, --output`：指定輸出的 Markdown 文件路徑
- `-t, --theme`：指定主題（default、gaia、uncover）
- `--pdf`：同時生成 PDF 文件
- `--pptx`：同時生成 PPTX 文件

示例：
```bash
python3 purepython-txt_to_marp_ppt.py input.txt -t gaia --pdf --pptx
```

## 功能說明

### 1. Markdown 轉換

- 自動將純文本轉換為 Marp 格式的 Markdown
- 支援三種精美主題：default、gaia、uncover
- 自動處理標題、列表和段落格式
- 提供美觀的默認 CSS 樣式
- 使用兩個空行分隔不同的幻燈片
- 自動將第一行文本設為幻燈片標題

### 2. PDF 轉換

- 將 Markdown 轉換為高質量 PDF 文件
- 保留所有樣式和格式
- 自動分頁
- 支援頁碼顯示

### 3. PPTX 轉換

- 生成可編輯的 PowerPoint 文件
- 完整保留主題樣式
- 支援所有 Office 軟件打開
- 方便後續編輯和修改

## 輸入文件格式說明

1. 使用純文本（.txt）文件作為輸入
2. 用兩個空行分隔不同的幻燈片
3. 每個幻燈片的第一行自動作為標題
4. 支援的格式：
   - 自動識別數字列表（1. 2. 3.）
   - 自動識別無序列表（- 或 *）
   - 支援 `<br>` 標籤換行
   - 保留單個空行作為段落分隔

## 注意事項

1. 確保已正確安裝 Marp CLI
2. 輸入文件必須是 UTF-8 編碼的文本文件
3. 轉換大文件時可能需要等待較長時間
4. 建議定期備份重要的簡報文件

## 故障排除

1. 如果出現「Marp CLI未安裝」錯誤：
   - 檢查 Node.js 和 npm 是否正確安裝
   - 重新執行 `npm install -g @marp-team/marp-cli`

2. 如果轉換失敗：
   - 檢查輸入文件編碼是否為 UTF-8
   - 確認是否有足夠的磁盤空間
   - 檢查輸出目錄的寫入權限

3. 如果輸出格式有問題：
   - 確認使用的是最新版本的 Marp CLI
   - 檢查輸入文件格式是否正確
   - 嘗試使用不同的主題