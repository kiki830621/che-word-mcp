# OOXML 概述與檔案結構

## 什麼是 OOXML？

Office Open XML (OOXML) 是 Microsoft Office 2007 及更新版本使用的文件格式標準。它是一個基於 XML 的開放標準，由 ECMA International 標準化為 ECMA-376，並被 ISO/IEC 採納為 ISO/IEC 29500。

### 關鍵特點

1. **基於 XML**：所有內容以 XML 格式儲存，可讀性高
2. **ZIP 封裝**：多個 XML 檔案打包成 ZIP 壓縮檔
3. **開放標準**：任何人都可以實作
4. **模組化設計**：不同功能分離到不同檔案

### 檔案副檔名

| 副檔名 | 說明 | MIME 類型 |
|--------|------|-----------|
| `.docx` | Word 文件 | `application/vnd.openxmlformats-officedocument.wordprocessingml.document` |
| `.docm` | 啟用巨集的 Word 文件 | `application/vnd.ms-word.document.macroEnabled.12` |
| `.dotx` | Word 範本 | `application/vnd.openxmlformats-officedocument.wordprocessingml.template` |
| `.xlsx` | Excel 活頁簿 | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| `.pptx` | PowerPoint 簡報 | `application/vnd.openxmlformats-officedocument.presentationml.presentation` |

---

## .docx 檔案結構

`.docx` 檔案本質上是一個 ZIP 壓縮檔，包含多個 XML 檔案和資源。

### 完整目錄結構

```
document.docx (ZIP 壓縮檔)
│
├── [Content_Types].xml              # 內容類型定義
│
├── _rels/
│   └── .rels                        # 套件層級關係
│
├── word/
│   ├── document.xml                 # ★ 主文件內容
│   ├── styles.xml                   # 樣式定義
│   ├── settings.xml                 # 文件設定
│   ├── webSettings.xml              # Web 設定
│   ├── fontTable.xml                # 字型表
│   ├── theme/
│   │   └── theme1.xml               # 佈景主題
│   ├── numbering.xml                # 編號/清單定義
│   ├── comments.xml                 # 註解
│   ├── commentsExtended.xml         # 擴展註解資訊
│   ├── footnotes.xml                # 腳註
│   ├── endnotes.xml                 # 尾註
│   ├── header1.xml                  # 頁首（可有多個）
│   ├── header2.xml
│   ├── footer1.xml                  # 頁尾（可有多個）
│   ├── footer2.xml
│   ├── glossary/                    # 建置組塊
│   │   └── document.xml
│   ├── media/                       # 嵌入媒體
│   │   ├── image1.png
│   │   ├── image2.jpeg
│   │   └── ...
│   ├── embeddings/                  # 嵌入物件
│   │   └── oleObject1.bin
│   └── _rels/
│       └── document.xml.rels        # 文件層級關係
│
├── docProps/
│   ├── core.xml                     # 核心屬性（標題、作者等）
│   ├── app.xml                      # 應用程式屬性
│   └── custom.xml                   # 自訂屬性
│
└── customXml/                       # 自訂 XML 資料
    ├── item1.xml
    └── itemProps1.xml
```

---

## 核心檔案詳解

### 1. [Content_Types].xml

定義套件中每個檔案的 MIME 類型。這是 OPC (Open Packaging Conventions) 規範的一部分。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <!-- 預設類型（依副檔名） -->
    <Default Extension="rels"
             ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml"
             ContentType="application/xml"/>
    <Default Extension="png"
             ContentType="image/png"/>
    <Default Extension="jpeg"
             ContentType="image/jpeg"/>

    <!-- 覆寫類型（依路徑） -->
    <Override PartName="/word/document.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/settings.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    <Override PartName="/word/fontTable.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
    <Override PartName="/word/numbering.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
    <Override PartName="/word/header1.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
    <Override PartName="/word/footer1.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
    <Override PartName="/word/comments.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
    <Override PartName="/word/footnotes.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
    <Override PartName="/word/endnotes.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
    <Override PartName="/docProps/core.xml"
              ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml"
              ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

### 2. _rels/.rels（套件關係）

定義套件根層級的關係。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <!-- 主文件 -->
    <Relationship Id="rId1"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                  Target="word/document.xml"/>
    <!-- 核心屬性 -->
    <Relationship Id="rId2"
                  Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                  Target="docProps/core.xml"/>
    <!-- 擴展屬性 -->
    <Relationship Id="rId3"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                  Target="docProps/app.xml"/>
</Relationships>
```

### 3. word/_rels/document.xml.rels（文件關係）

定義主文件與其他部件的關係。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <!-- 樣式 -->
    <Relationship Id="rId1"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                  Target="styles.xml"/>
    <!-- 設定 -->
    <Relationship Id="rId2"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
                  Target="settings.xml"/>
    <!-- 字型表 -->
    <Relationship Id="rId3"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
                  Target="fontTable.xml"/>
    <!-- 編號定義 -->
    <Relationship Id="rId4"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
                  Target="numbering.xml"/>
    <!-- 頁首 -->
    <Relationship Id="rId5"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
                  Target="header1.xml"/>
    <!-- 頁尾 -->
    <Relationship Id="rId6"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
                  Target="footer1.xml"/>
    <!-- 圖片 -->
    <Relationship Id="rId7"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                  Target="media/image1.png"/>
    <!-- 外部超連結 -->
    <Relationship Id="rId8"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                  Target="https://example.com"
                  TargetMode="External"/>
    <!-- 註解 -->
    <Relationship Id="rId9"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
                  Target="comments.xml"/>
    <!-- 腳註 -->
    <Relationship Id="rId10"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
                  Target="footnotes.xml"/>
</Relationships>
```

### 4. docProps/core.xml（核心屬性）

Dublin Core 元資料。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">

    <dc:title>文件標題</dc:title>
    <dc:subject>主題</dc:subject>
    <dc:creator>作者名稱</dc:creator>
    <cp:keywords>關鍵字1, 關鍵字2</cp:keywords>
    <dc:description>文件描述</dc:description>
    <cp:lastModifiedBy>最後修改者</cp:lastModifiedBy>
    <cp:revision>1</cp:revision>
    <dcterms:created xsi:type="dcterms:W3CDTF">2025-01-13T10:00:00Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-13T12:00:00Z</dcterms:modified>
    <cp:category>類別</cp:category>
    <cp:contentStatus>草稿</cp:contentStatus>
</cp:coreProperties>
```

### 5. docProps/app.xml（應用程式屬性）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">

    <Application>Microsoft Office Word</Application>
    <AppVersion>16.0000</AppVersion>
    <DocSecurity>0</DocSecurity>
    <Template>Normal.dotm</Template>
    <TotalTime>60</TotalTime>
    <Pages>10</Pages>
    <Words>2500</Words>
    <Characters>15000</Characters>
    <CharactersWithSpaces>17500</CharactersWithSpaces>
    <Paragraphs>50</Paragraphs>
    <Lines>200</Lines>
    <Company>公司名稱</Company>
    <Manager>經理名稱</Manager>

    <!-- 標題資訊 -->
    <HeadingPairs>
        <vt:vector size="2" baseType="variant">
            <vt:variant><vt:lpstr>標題</vt:lpstr></vt:variant>
            <vt:variant><vt:i4>3</vt:i4></vt:variant>
        </vt:vector>
    </HeadingPairs>
    <TitlesOfParts>
        <vt:vector size="3" baseType="lpstr">
            <vt:lpstr>第一章</vt:lpstr>
            <vt:lpstr>第二章</vt:lpstr>
            <vt:lpstr>第三章</vt:lpstr>
        </vt:vector>
    </TitlesOfParts>
</Properties>
```

---

## 最小可行的 .docx 檔案

建立一個有效的 .docx 檔案最少需要以下檔案：

```
minimal.docx
├── [Content_Types].xml
├── _rels/
│   └── .rels
└── word/
    └── document.xml
```

### 最小 [Content_Types].xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
```

### 最小 _rels/.rels

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
```

### 最小 word/document.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>Hello, World!</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>
```

---

## OPC (Open Packaging Conventions)

OOXML 建立在 OPC 規範之上，OPC 定義了：

### 核心概念

| 概念 | 說明 |
|------|------|
| **Package** | ZIP 容器，包含所有部件 |
| **Part** | 套件中的單個檔案 |
| **Content Type** | 每個部件的 MIME 類型 |
| **Relationship** | 部件之間的連結 |

### 關係類型

關係分為兩種：

1. **內部關係**（`TargetMode="Internal"` 或省略）
   - 指向套件內的其他部件
   - Target 是相對路徑

2. **外部關係**（`TargetMode="External"`）
   - 指向套件外的資源
   - Target 是絕對 URI

---

## 使用程式解壓縮 .docx

### Python

```python
import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile('document.docx', 'r') as zip_ref:
    # 讀取主文件
    with zip_ref.open('word/document.xml') as f:
        tree = ET.parse(f)
        root = tree.getroot()
```

### Swift

```swift
import ZIPFoundation

let fileURL = URL(fileURLWithPath: "document.docx")
guard let archive = Archive(url: fileURL, accessMode: .read) else { return }

if let entry = archive["word/document.xml"] {
    var data = Data()
    _ = try archive.extract(entry) { data.append($0) }
    // 解析 XML
}
```

### 命令列

```bash
# 解壓縮
unzip document.docx -d extracted/

# 格式化查看 XML
xmllint --format extracted/word/document.xml

# 重新打包
cd extracted && zip -r ../new_document.docx . && cd ..
```

---

## 下一步

- [02-namespaces.md](02-namespaces.md) - 深入了解 XML 命名空間
- [10-document.md](10-document.md) - 學習 document.xml 的結構
