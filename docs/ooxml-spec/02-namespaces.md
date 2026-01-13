# XML 命名空間參考

## 主要命名空間

OOXML 使用多個 XML 命名空間來區分不同功能模組。

### WordprocessingML 核心命名空間

| 前綴 | 命名空間 URI | 用途 |
|------|-------------|------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | 主要文件元素 |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` | Word 2010 擴展 |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` | Word 2013 擴展 |
| `w16` | `http://schemas.microsoft.com/office/word/2018/wordml` | Word 2019 擴展 |
| `wpc` | `http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas` | 繪圖畫布 |
| `wpg` | `http://schemas.microsoft.com/office/word/2010/wordprocessingGroup` | 圖形群組 |
| `wps` | `http://schemas.microsoft.com/office/word/2010/wordprocessingShape` | 圖形 |

### 關係命名空間

| 前綴 | 命名空間 URI | 用途 |
|------|-------------|------|
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | 關係參照 |
| `pr` | `http://schemas.openxmlformats.org/package/2006/relationships` | 套件關係 |

### DrawingML 命名空間

| 前綴 | 命名空間 URI | 用途 |
|------|-------------|------|
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML 核心 |
| `pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` | 圖片 |
| `wp` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | Word 圖形定位 |
| `a14` | `http://schemas.microsoft.com/office/drawing/2010/main` | DrawingML 2010 擴展 |

### 數學公式命名空間

| 前綴 | 命名空間 URI | 用途 |
|------|-------------|------|
| `m` | `http://schemas.openxmlformats.org/officeDocument/2006/math` | Office Math (OMML) |

### 其他命名空間

| 前綴 | 命名空間 URI | 用途 |
|------|-------------|------|
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | 標記相容性 |
| `o` | `urn:schemas-microsoft-com:office:office` | 舊版 Office |
| `v` | `urn:schemas-microsoft-com:vml` | VML 向量圖形 |
| `ve` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | 版本擴展 |
| `sl` | `http://schemas.openxmlformats.org/schemaLibrary/2006/main` | Schema 庫 |

---

## 命名空間聲明範例

### document.xml 標準聲明

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    mc:Ignorable="w14 w15">
    <w:body>
        <!-- 文件內容 -->
    </w:body>
</w:document>
```

### styles.xml 標準聲明

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    mc:Ignorable="w14">
    <!-- 樣式定義 -->
</w:styles>
```

---

## 常用元素與其命名空間

### w: 命名空間元素

| 元素 | 說明 |
|------|------|
| `w:document` | 文件根元素 |
| `w:body` | 文件主體 |
| `w:p` | 段落 (Paragraph) |
| `w:pPr` | 段落屬性 (Paragraph Properties) |
| `w:r` | 文字運行 (Run) |
| `w:rPr` | 文字運行屬性 (Run Properties) |
| `w:t` | 文字 (Text) |
| `w:tbl` | 表格 (Table) |
| `w:tr` | 表格列 (Table Row) |
| `w:tc` | 表格儲存格 (Table Cell) |
| `w:sectPr` | 分節屬性 (Section Properties) |
| `w:style` | 樣式定義 |
| `w:hyperlink` | 超連結 |
| `w:bookmarkStart` | 書籤開始 |
| `w:bookmarkEnd` | 書籤結束 |
| `w:comment` | 註解 |
| `w:footnote` | 腳註 |
| `w:endnote` | 尾註 |

### r: 命名空間屬性

| 屬性 | 說明 |
|------|------|
| `r:id` | 關係 ID 參照 |
| `r:embed` | 嵌入資源 ID |
| `r:link` | 連結資源 ID |

### wp: 命名空間元素

| 元素 | 說明 |
|------|------|
| `wp:inline` | 行內定位圖形 |
| `wp:anchor` | 浮動定位圖形 |
| `wp:extent` | 圖形尺寸 |
| `wp:docPr` | 圖形屬性 |
| `wp:cNvGraphicFramePr` | 非視覺圖形框架屬性 |

### a: 命名空間元素

| 元素 | 說明 |
|------|------|
| `a:graphic` | 圖形容器 |
| `a:graphicData` | 圖形資料 |
| `a:blip` | 圖片參照 |
| `a:stretch` | 延展設定 |
| `a:fillRect` | 填充矩形 |

### pic: 命名空間元素

| 元素 | 說明 |
|------|------|
| `pic:pic` | 圖片元素 |
| `pic:nvPicPr` | 非視覺圖片屬性 |
| `pic:blipFill` | 圖片填充 |
| `pic:spPr` | 形狀屬性 |

### m: 命名空間元素

| 元素 | 說明 |
|------|------|
| `m:oMath` | 數學公式區塊 |
| `m:oMathPara` | 數學段落 |
| `m:r` | 數學運行 |
| `m:t` | 數學文字 |
| `m:f` | 分數 |
| `m:rad` | 根號 |
| `m:sup` | 上標 |
| `m:sub` | 下標 |

---

## 標記相容性 (Markup Compatibility)

### mc:Ignorable

允許處理器忽略不認識的命名空間：

```xml
<w:document
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    mc:Ignorable="w14 w15">
```

### mc:AlternateContent

提供向後相容的替代內容：

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice Requires="w14">
        <!-- Word 2010+ 專用內容 -->
        <w14:contentPart r:id="rId1"/>
    </mc:Choice>
    <mc:Fallback>
        <!-- 舊版本替代內容 -->
        <w:p>
            <w:r>
                <w:t>[此內容需要 Word 2010 或更新版本]</w:t>
            </w:r>
        </w:p>
    </mc:Fallback>
</mc:AlternateContent>
```

---

## 關係類型 URI

### 文件關係類型

| 關係類型 | URI |
|----------|-----|
| 主文件 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument` |
| 樣式 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles` |
| 設定 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings` |
| 字型表 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable` |
| 編號 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering` |
| 頁首 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/header` |
| 頁尾 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer` |
| 圖片 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image` |
| 超連結 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink` |
| 註解 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` |
| 腳註 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes` |
| 尾註 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes` |
| 佈景主題 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme` |

### 套件關係類型

| 關係類型 | URI |
|----------|-----|
| 核心屬性 | `http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties` |
| 擴展屬性 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties` |
| 縮圖 | `http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail` |

---

## 下一步

- [03-content-types.md](03-content-types.md) - Content Types 詳解
- [10-document.md](10-document.md) - document.xml 結構
