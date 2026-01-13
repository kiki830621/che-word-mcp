# 頁首與頁尾 (Headers & Footers)

## 概述

頁首和頁尾存儲在獨立的 XML 檔案中（如 `header1.xml`、`footer1.xml`），透過關係（relationships）連結到主文件。

## 檔案結構

```
word/
├── document.xml          # 主文件
├── header1.xml           # 頁首（預設/奇數頁）
├── header2.xml           # 頁首（首頁）
├── header3.xml           # 頁首（偶數頁）
├── footer1.xml           # 頁尾（預設/奇數頁）
├── footer2.xml           # 頁尾（首頁）
├── footer3.xml           # 頁尾（偶數頁）
└── _rels/
    └── document.xml.rels # 關係定義
```

---

## 關係設定

### document.xml.rels

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
                  Target="header1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
                  Target="header2.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
                  Target="footer1.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
                  Target="footer2.xml"/>
</Relationships>
```

### [Content_Types].xml

```xml
<Override PartName="/word/header1.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/header2.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/word/footer2.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
```

### sectPr 參照

```xml
<w:sectPr>
    <w:headerReference w:type="default" r:id="rId1"/>
    <w:headerReference w:type="first" r:id="rId2"/>
    <w:footerReference w:type="default" r:id="rId3"/>
    <w:footerReference w:type="first" r:id="rId4"/>
    <w:titlePg/>  <!-- 啟用首頁不同 -->
</w:sectPr>
```

---

## 頁首/頁尾類型

| w:type | 說明 | 使用情境 |
|--------|------|----------|
| `default` | 預設（奇數頁） | 一般頁面 |
| `first` | 首頁 | 封面或目錄頁 |
| `even` | 偶數頁 | 雙面印刷 |

### 設定組合

| 需求 | 設定 |
|------|------|
| 所有頁面相同 | 只設定 `default` |
| 首頁不同 | 設定 `default` + `first`，加上 `<w:titlePg/>` |
| 奇偶頁不同 | 設定 `default` + `even`，加上 `<w:evenAndOddHeaders/>` |
| 首頁+奇偶頁不同 | 三者都設定，加上 `<w:titlePg/>` 和 `<w:evenAndOddHeaders/>` |

---

## w:hdr（頁首結構）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:p>
        <w:pPr>
            <w:pStyle w:val="Header"/>
        </w:pPr>
        <w:r>
            <w:t>頁首文字</w:t>
        </w:r>
    </w:p>
</w:hdr>
```

### w:hdr 子元素

| 元素 | 說明 |
|------|------|
| `w:p` | 段落 |
| `w:tbl` | 表格 |
| `w:sdt` | 結構化文件標籤 |
| `w:customXml` | 自訂 XML |
| `w:altChunk` | 替代內容區塊 |

---

## w:ftr（頁尾結構）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:p>
        <w:pPr>
            <w:pStyle w:val="Footer"/>
            <w:jc w:val="center"/>
        </w:pPr>
        <!-- 頁碼欄位 -->
        <w:r>
            <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
            <w:instrText xml:space="preserve"> PAGE </w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
            <w:t>1</w:t>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end"/>
        </w:r>
    </w:p>
</w:ftr>
```

---

## 頁碼欄位

### 簡單頁碼

```xml
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> PAGE </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>1</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### 頁碼格式

| 欄位代碼 | 說明 | 輸出範例 |
|----------|------|----------|
| `PAGE` | 目前頁碼 | 1 |
| `NUMPAGES` | 總頁數 | 10 |
| `PAGE \* MERGEFORMAT` | 保留格式 | 1 |
| `PAGE \* Roman` | 羅馬數字 | I |
| `PAGE \* Arabic` | 阿拉伯數字 | 1 |

### 「第 X 頁，共 Y 頁」格式

```xml
<w:p>
    <w:pPr>
        <w:pStyle w:val="Footer"/>
        <w:jc w:val="center"/>
    </w:pPr>
    <!-- 前綴文字 -->
    <w:r>
        <w:t xml:space="preserve">第 </w:t>
    </w:r>
    <!-- PAGE 欄位 -->
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText xml:space="preserve"> PAGE </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
        <w:t>1</w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
    <!-- 中間文字 -->
    <w:r>
        <w:t xml:space="preserve"> 頁，共 </w:t>
    </w:r>
    <!-- NUMPAGES 欄位 -->
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText xml:space="preserve"> NUMPAGES </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
        <w:t>10</w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
    <!-- 後綴文字 -->
    <w:r>
        <w:t xml:space="preserve"> 頁</w:t>
    </w:r>
</w:p>
```

### "Page X of Y" 格式

```xml
<w:p>
    <w:pPr>
        <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
        <w:t xml:space="preserve">Page </w:t>
    </w:r>
    <!-- PAGE -->
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r>
        <w:t xml:space="preserve"> of </w:t>
    </w:r>
    <!-- NUMPAGES -->
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> NUMPAGES </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>10</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
```

---

## 常用頁首/頁尾內容

### 文件標題

```xml
<w:p>
    <w:pPr><w:pStyle w:val="Header"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> TITLE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>文件標題</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
```

### 作者

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> AUTHOR </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>作者名稱</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

### 日期

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> DATE \@ "yyyy/MM/dd" </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>2025/01/13</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

### 檔案名稱

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> FILENAME </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>document.docx</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

---

## 使用表格排版

### 三欄頁首

```xml
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:tbl>
        <w:tblPr>
            <w:tblW w:w="5000" w:type="pct"/>
            <w:tblBorders>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            </w:tblBorders>
        </w:tblPr>
        <w:tblGrid>
            <w:gridCol w:w="3000"/>
            <w:gridCol w:w="3000"/>
            <w:gridCol w:w="3000"/>
        </w:tblGrid>
        <w:tr>
            <!-- 左：公司名稱 -->
            <w:tc>
                <w:tcPr>
                    <w:tcBorders>
                        <w:top w:val="nil"/>
                        <w:left w:val="nil"/>
                        <w:bottom w:val="nil"/>
                        <w:right w:val="nil"/>
                    </w:tcBorders>
                </w:tcPr>
                <w:p>
                    <w:pPr><w:jc w:val="left"/></w:pPr>
                    <w:r><w:t>公司名稱</w:t></w:r>
                </w:p>
            </w:tc>
            <!-- 中：文件標題 -->
            <w:tc>
                <w:tcPr>
                    <w:tcBorders>
                        <w:top w:val="nil"/>
                        <w:left w:val="nil"/>
                        <w:bottom w:val="nil"/>
                        <w:right w:val="nil"/>
                    </w:tcBorders>
                </w:tcPr>
                <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:rPr><w:b/></w:rPr><w:t>文件標題</w:t></w:r>
                </w:p>
            </w:tc>
            <!-- 右：日期 -->
            <w:tc>
                <w:tcPr>
                    <w:tcBorders>
                        <w:top w:val="nil"/>
                        <w:left w:val="nil"/>
                        <w:bottom w:val="nil"/>
                        <w:right w:val="nil"/>
                    </w:tcBorders>
                </w:tcPr>
                <w:p>
                    <w:pPr><w:jc w:val="right"/></w:pPr>
                    <w:r><w:t>2025/01/13</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
    </w:tbl>
</w:hdr>
```

---

## 完整範例

### header1.xml（預設頁首）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:p>
        <w:pPr>
            <w:pStyle w:val="Header"/>
            <w:pBdr>
                <w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>
            </w:pBdr>
            <w:tabs>
                <w:tab w:val="center" w:pos="4680"/>
                <w:tab w:val="right" w:pos="9360"/>
            </w:tabs>
        </w:pPr>
        <w:r>
            <w:t>公司機密</w:t>
        </w:r>
        <w:r>
            <w:tab/>
        </w:r>
        <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>技術規格文件</w:t>
        </w:r>
        <w:r>
            <w:tab/>
        </w:r>
        <w:r>
            <w:t>v1.0</w:t>
        </w:r>
    </w:p>
</w:hdr>
```

### footer1.xml（預設頁尾）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:p>
        <w:pPr>
            <w:pStyle w:val="Footer"/>
            <w:pBdr>
                <w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>
            </w:pBdr>
            <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
            <w:t xml:space="preserve">第 </w:t>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
            <w:instrText xml:space="preserve"> PAGE </w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
            <w:rPr><w:noProof/></w:rPr>
            <w:t>1</w:t>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end"/>
        </w:r>
        <w:r>
            <w:t xml:space="preserve"> 頁，共 </w:t>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
            <w:instrText xml:space="preserve"> NUMPAGES </w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
            <w:rPr><w:noProof/></w:rPr>
            <w:t>1</w:t>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end"/>
        </w:r>
        <w:r>
            <w:t xml:space="preserve"> 頁</w:t>
        </w:r>
    </w:p>
</w:ftr>
```

### header2.xml（首頁頁首 - 空白）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:p>
        <w:pPr>
            <w:pStyle w:val="Header"/>
        </w:pPr>
    </w:p>
</w:hdr>
```

---

## 下一步

- [43-page-numbers.md](43-page-numbers.md) - 頁碼設定詳解
- [50-images.md](50-images.md) - 圖片
- [64-fields.md](64-fields.md) - 欄位代碼
