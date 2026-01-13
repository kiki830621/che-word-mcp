# OOXML WordprocessingML 完整規範文件

本文件集完整涵蓋 Office Open XML (OOXML) WordprocessingML 規範，基於 ECMA-376 標準，專為 che-word-mcp 開發參考。

## 目錄

### 基礎篇

| 文件 | 說明 | 狀態 |
|------|------|------|
| [01-overview.md](01-overview.md) | OOXML 概述與檔案結構 | ✅ |
| [02-namespaces.md](02-namespaces.md) | XML 命名空間參考 | ✅ |

### 文件結構篇

| 文件 | 說明 | 狀態 |
|------|------|------|
| [10-document.md](10-document.md) | document.xml 主文件結構 | ✅ |
| [11-paragraph.md](11-paragraph.md) | 段落 (Paragraph) 元素 | ✅ |
| [12-run.md](12-run.md) | 文字運行 (Run) 元素 | ✅ |
| [13-text-formatting.md](13-text-formatting.md) | 文字格式化屬性 | ✅ |
| [14-paragraph-formatting.md](14-paragraph-formatting.md) | 段落格式化屬性 | ✅ |

### 表格篇

| 文件 | 說明 | 狀態 |
|------|------|------|
| [20-table.md](20-table.md) | 表格 (Table) 結構 | ✅ |
| [21-table-row.md](21-table-row.md) | 表格列 (Row) 元素 | ✅ |
| [22-table-cell.md](22-table-cell.md) | 表格儲存格 (Cell) 元素 | ✅ |

### 樣式篇

| 文件 | 說明 | 狀態 |
|------|------|------|
| [30-styles.md](30-styles.md) | styles.xml 樣式系統 | ✅ |
| [32-numbering.md](32-numbering.md) | 編號與清單定義 | ✅ |

### 頁面設定篇

| 文件 | 說明 | 狀態 |
|------|------|------|
| [40-section.md](40-section.md) | 分節 (Section) 屬性 | ✅ |
| [42-headers-footers.md](42-headers-footers.md) | 頁首與頁尾 | ✅ |

### 媒體與連結篇

| 文件 | 說明 | 狀態 |
|------|------|------|
| [50-images.md](50-images.md) | 圖片與 DrawingML | ✅ |
| [51-hyperlinks.md](51-hyperlinks.md) | 超連結與書籤 | ✅ |

### 進階功能篇

| 文件 | 說明 | 狀態 |
|------|------|------|
| [60-comments.md](60-comments.md) | 註解 (Comments) | ✅ |
| [61-track-changes.md](61-track-changes.md) | 追蹤修訂 | ✅ |
| [62-footnotes-endnotes.md](62-footnotes-endnotes.md) | 腳註與尾註 | ✅ |
| [63-toc.md](63-toc.md) | 目錄 (Table of Contents) | ✅ |
| [64-fields.md](64-fields.md) | 欄位代碼 (Field Codes) | ✅ |
| [65-forms.md](65-forms.md) | 表單控制項 | ✅ |
| [66-math.md](66-math.md) | 數學公式 (OMML) | ✅ |

### 附錄

| 文件 | 說明 | 狀態 |
|------|------|------|
| [A1-units.md](A1-units.md) | 度量單位換算 | ✅ |
| [A2-colors.md](A2-colors.md) | 顏色表示法 | ✅ |
| [A3-fonts.md](A3-fonts.md) | 字型處理 | ✅ |
| [A4-compatibility.md](A4-compatibility.md) | 相容性設定 | ✅ |

---

## 文件統計

- **總文件數**：28 篇
- **基礎篇**：2 篇
- **文件結構篇**：5 篇
- **表格篇**：3 篇
- **樣式篇**：2 篇
- **頁面設定篇**：2 篇
- **媒體與連結篇**：2 篇
- **進階功能篇**：7 篇
- **附錄**：4 篇

---

## 參考資源

- [ECMA-376 官方規範](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [ISO/IEC 29500](https://www.iso.org/standard/71691.html)
- [Microsoft Open Specifications](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-docx/)

## 版本資訊

- OOXML 版本：ECMA-376 5th Edition (2015)
- 文件版本：1.0.0
- 最後更新：2026-01-13
- 維護者：che-word-mcp 專案

---

## 快速參考

### .docx 檔案結構

```
document.docx (ZIP)
├── [Content_Types].xml          # MIME 類型定義
├── _rels/
│   └── .rels                    # 套件關係
├── word/
│   ├── document.xml             # 主文件內容
│   ├── styles.xml               # 樣式定義
│   ├── settings.xml             # 文件設定
│   ├── fontTable.xml            # 字型表
│   ├── numbering.xml            # 編號定義
│   ├── comments.xml             # 註解
│   ├── footnotes.xml            # 腳註
│   ├── endnotes.xml             # 尾註
│   ├── header1.xml              # 頁首
│   ├── footer1.xml              # 頁尾
│   ├── media/                   # 嵌入媒體
│   │   └── image1.png
│   └── _rels/
│       └── document.xml.rels    # 文件關係
└── docProps/
    ├── core.xml                 # 核心屬性
    └── app.xml                  # 應用程式屬性
```

### 常用命名空間前綴

| 前綴 | 命名空間 | 用途 |
|------|----------|------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | WordprocessingML 主要元素 |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | 關係參照 |
| `wp` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | 圖形定位 |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML |
| `pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` | 圖片 |
| `m` | `http://schemas.openxmlformats.org/officeDocument/2006/math` | Office Math |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` | Word 2010 擴展 |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` | Word 2012 擴展 |

### 常用單位換算

| 單位 | 說明 | 1 英吋 = |
|------|------|----------|
| Twips | 1/20 點 | 1440 twips |
| Half-points | 字型大小單位 | - |
| EMU | 圖片尺寸單位 | 914,400 EMU |
| Pct (1/50%) | 百分比 | 5000 = 100% |

### 快速文件結構

```xml
<w:document>
    <w:body>
        <w:p>                    <!-- 段落 -->
            <w:pPr>              <!-- 段落屬性 -->
                <w:pStyle/>
                <w:jc/>
            </w:pPr>
            <w:r>                <!-- Run -->
                <w:rPr>          <!-- Run 屬性 -->
                    <w:b/>
                    <w:sz/>
                </w:rPr>
                <w:t>文字</w:t>  <!-- 文字 -->
            </w:r>
        </w:p>
        <w:tbl>                  <!-- 表格 -->
            <w:tblPr/>           <!-- 表格屬性 -->
            <w:tr>               <!-- 列 -->
                <w:tc>           <!-- 儲存格 -->
                    <w:p/>
                </w:tc>
            </w:tr>
        </w:tbl>
        <w:sectPr>               <!-- 分節屬性 -->
            <w:pgSz/>
            <w:pgMar/>
        </w:sectPr>
    </w:body>
</w:document>
```
