# che-word-mcp

以 Swift 原生開發的 MCP (Model Context Protocol) 伺服器，用於操作 Microsoft Word 文件 (.docx)。這是**首個 Swift OOXML 函式庫**，直接操作 Office Open XML 而不依賴任何第三方 Word 函式庫。

[English](README.md)

## 特色

- **純 Swift 實作**：不需要 Node.js、Python 或其他執行環境
- **直接操作 OOXML**：直接處理 XML，不需要安裝 Microsoft Word
- **單一執行檔**：只有一個 binary 檔案
- **18 個 MCP 工具**：完整的文件操作功能
- **跨平台**：支援 macOS（以及其他支援 Swift 的平台）

## 安裝

### 系統需求

- macOS 13.0+ (Ventura 或更新版本)
- Swift 5.9+

### 從原始碼編譯

```bash
git clone https://github.com/kiki830621/che-word-mcp.git
cd che-word-mcp
swift build -c release
```

執行檔位於 `.build/release/CheWordMCP`

### 加入 Claude Code

```bash
claude mcp add che-word-mcp /path/to/che-word-mcp/.build/release/CheWordMCP
```

### 加入 Claude Desktop

編輯 `~/Library/Application Support/Claude/claude_desktop_config.json`：

```json
{
  "mcpServers": {
    "che-word-mcp": {
      "command": "/path/to/che-word-mcp/.build/release/CheWordMCP"
    }
  }
}
```

## 可用工具

### 文件管理 (6 個工具)

| 工具 | 說明 |
|------|------|
| `create_document` | 建立新的 Word 文件 |
| `open_document` | 開啟現有的 .docx 檔案 |
| `save_document` | 儲存文件為 .docx |
| `close_document` | 關閉已開啟的文件 |
| `list_open_documents` | 列出所有已開啟的文件 |
| `get_document_info` | 取得文件統計資訊（段落數、字數、字元數）|

### 內容操作 (6 個工具)

| 工具 | 說明 |
|------|------|
| `get_text` | 取得純文字內容 |
| `get_paragraphs` | 取得所有段落（含格式資訊）|
| `insert_paragraph` | 插入新段落 |
| `update_paragraph` | 更新現有段落內容 |
| `delete_paragraph` | 刪除段落 |
| `replace_text` | 搜尋並取代文字 |

### 格式化 (3 個工具)

| 工具 | 說明 |
|------|------|
| `format_text` | 設定文字格式（粗體、斜體、顏色、字型）|
| `set_paragraph_format` | 設定段落格式（對齊、間距）|
| `apply_style` | 套用內建樣式（Heading1、Title 等）|

### 表格 (1 個工具)

| 工具 | 說明 |
|------|------|
| `insert_table` | 插入表格（可含資料）|

### 匯出 (2 個工具)

| 工具 | 說明 |
|------|------|
| `export_text` | 匯出為純文字 |
| `export_markdown` | 匯出為 Markdown |

## 使用範例

### 建立含標題和內文的文件

```
建立一個新的 Word 文件叫做 "report"：
- 標題：「季度報告」
- 一級標題：「簡介」
- 一段說明報告目的的文字
儲存到 ~/Documents/report.docx
```

### 開啟並修改現有文件

```
開啟 ~/Documents/proposal.docx
將所有的 "2024" 取代為 "2025"
儲存變更
```

### 建立含表格的文件

```
建立一個文件，包含一個 3x4 的表格：
- 表頭：姓名、年齡、部門
- 資料列填入員工資訊
同時匯出為 Markdown
```

## 技術細節

### OOXML 結構

伺服器產生符合標準的 Office Open XML 文件，結構如下：

```
document.docx (ZIP)
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── word/
│   ├── document.xml      # 主要內容
│   ├── styles.xml        # 樣式定義
│   ├── settings.xml      # 文件設定
│   ├── fontTable.xml     # 字型定義
│   └── _rels/
│       └── document.xml.rels
└── docProps/
    ├── core.xml          # 中繼資料（作者、標題）
    └── app.xml           # 應用程式資訊
```

### 依賴套件

- [MCP Swift SDK](https://github.com/modelcontextprotocol/swift-sdk) (v0.10.0+) - Model Context Protocol 實作
- [ZIPFoundation](https://github.com/weichsel/ZIPFoundation) (v0.9.0+) - ZIP 壓縮/解壓縮

## 與其他方案比較

| 功能 | Anthropic Word MCP | python-docx | docx npm | **che-word-mcp** |
|------|-------------------|-------------|----------|------------------|
| 語言 | Node.js | Python | Node.js | **Swift** |
| 後端 | AppleScript | OOXML | OOXML | **OOXML** |
| 需要 Word | 是 | 否 | 否 | **否** |
| 執行環境 | Node.js | Python | Node.js | **無** |
| 單一執行檔 | 否 | 否 | 否 | **是** |

## 開發計畫

- [ ] 圖片支援
- [ ] 頁首/頁尾支援
- [ ] 分頁和分節
- [ ] 編號/項目符號清單
- [ ] 註解和追蹤修訂
- [ ] 讀取現有文件樣式

## 授權

MIT License

## 作者

鄭澈 ([@kiki830621](https://github.com/kiki830621))

## 相關專案

- [che-apple-mail-mcp](https://github.com/kiki830621/che-apple-mail-mcp) - Apple Mail MCP 伺服器
- [che-ical-mcp](https://github.com/kiki830621/che-ical-mcp) - macOS 行事曆 MCP 伺服器
