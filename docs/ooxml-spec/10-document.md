# document.xml 主文件結構

## 概述

`word/document.xml` 是 .docx 檔案的核心，包含文件的所有內容。

## 基本結構

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <!-- 文件內容 -->
        <w:p>...</w:p>           <!-- 段落 -->
        <w:tbl>...</w:tbl>       <!-- 表格 -->
        <w:sdt>...</w:sdt>       <!-- 結構化文件標籤 -->

        <!-- 分節屬性（最後一節） -->
        <w:sectPr>...</w:sectPr>
    </w:body>
</w:document>
```

---

## w:document 元素

根元素，包含整個文件。

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:conformance` | 一致性等級：`strict` 或 `transitional` |

### 子元素

| 元素 | 說明 | 必要 |
|------|------|------|
| `w:body` | 文件主體 | 是 |
| `w:background` | 文件背景 | 否 |

---

## w:body 元素

文件主體，包含所有內容。

### 子元素（Block-Level）

| 元素 | 說明 |
|------|------|
| `w:p` | 段落 |
| `w:tbl` | 表格 |
| `w:sdt` | 結構化文件標籤 (Content Control) |
| `w:customXml` | 自訂 XML 區塊 |
| `w:sectPr` | 分節屬性（僅最後一個） |

### 內容模型

```
w:body = (
    (w:p | w:tbl | w:sdt | w:customXml | w:altChunk)*,
    w:sectPr?
)
```

---

## 完整範例

### 包含多種元素的文件

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <w:body>

        <!-- 標題段落 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Title"/>
            </w:pPr>
            <w:r>
                <w:t>文件標題</w:t>
            </w:r>
        </w:p>

        <!-- 一級標題 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:r>
                <w:t>第一章 簡介</w:t>
            </w:r>
        </w:p>

        <!-- 一般段落 -->
        <w:p>
            <w:r>
                <w:t>這是一段普通的文字。</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:b/>
                </w:rPr>
                <w:t>這是粗體文字。</w:t>
            </w:r>
        </w:p>

        <!-- 含超連結的段落 -->
        <w:p>
            <w:r>
                <w:t>請訪問 </w:t>
            </w:r>
            <w:hyperlink r:id="rId5">
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="Hyperlink"/>
                    </w:rPr>
                    <w:t>我們的網站</w:t>
                </w:r>
            </w:hyperlink>
            <w:r>
                <w:t> 了解更多。</w:t>
            </w:r>
        </w:p>

        <!-- 項目符號清單 -->
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="1"/>
                </w:numPr>
            </w:pPr>
            <w:r>
                <w:t>第一個項目</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="1"/>
                </w:numPr>
            </w:pPr>
            <w:r>
                <w:t>第二個項目</w:t>
            </w:r>
        </w:p>

        <!-- 表格 -->
        <w:tbl>
            <w:tblPr>
                <w:tblStyle w:val="TableGrid"/>
                <w:tblW w:w="5000" w:type="pct"/>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="4500"/>
                <w:gridCol w:w="4500"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr>
                        <w:shd w:val="clear" w:fill="CCCCCC"/>
                    </w:tcPr>
                    <w:p>
                        <w:r><w:t>標題 A</w:t></w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr>
                        <w:shd w:val="clear" w:fill="CCCCCC"/>
                    </w:tcPr>
                    <w:p>
                        <w:r><w:t>標題 B</w:t></w:r>
                    </w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r><w:t>資料 1</w:t></w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r><w:t>資料 2</w:t></w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>

        <!-- 含圖片的段落 -->
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="1905000" cy="1428750"/>
                        <wp:docPr id="1" name="Picture 1"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic>
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="1" name="image1.png"/>
                                        <pic:cNvPicPr/>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="rId6"/>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="1905000" cy="1428750"/>
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst/>
                                        </a:prstGeom>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>

        <!-- 分頁符 -->
        <w:p>
            <w:r>
                <w:br w:type="page"/>
            </w:r>
        </w:p>

        <!-- 第二頁內容 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:r>
                <w:t>第二章 詳細內容</w:t>
            </w:r>
        </w:p>

        <!-- 含註解的段落 -->
        <w:p>
            <w:r>
                <w:t>這段文字有</w:t>
            </w:r>
            <w:commentRangeStart w:id="0"/>
            <w:r>
                <w:t>註解</w:t>
            </w:r>
            <w:commentRangeEnd w:id="0"/>
            <w:r>
                <w:commentReference w:id="0"/>
            </w:r>
            <w:r>
                <w:t>標記。</w:t>
            </w:r>
        </w:p>

        <!-- 含腳註的段落 -->
        <w:p>
            <w:r>
                <w:t>這是需要參考資料的內容</w:t>
            </w:r>
            <w:r>
                <w:footnoteReference w:id="1"/>
            </w:r>
            <w:r>
                <w:t>。</w:t>
            </w:r>
        </w:p>

        <!-- 分節屬性 -->
        <w:sectPr>
            <!-- 頁面大小 (A4) -->
            <w:pgSz w:w="11906" w:h="16838"/>
            <!-- 頁邊距 -->
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                     w:header="720" w:footer="720" w:gutter="0"/>
            <!-- 頁首參照 -->
            <w:headerReference w:type="default" r:id="rId7"/>
            <!-- 頁尾參照 -->
            <w:footerReference w:type="default" r:id="rId8"/>
            <!-- 欄設定 -->
            <w:cols w:space="720"/>
            <!-- 文件格線 -->
            <w:docGrid w:linePitch="360"/>
        </w:sectPr>

    </w:body>
</w:document>
```

---

## Block-Level 元素詳解

### w:p（段落）

段落是文件的基本組成單位。

```xml
<w:p>
    <w:pPr>...</w:pPr>    <!-- 段落屬性 -->
    <w:r>...</w:r>         <!-- 文字運行 -->
    <w:hyperlink>...</w:hyperlink>  <!-- 超連結 -->
    <w:bookmarkStart/>     <!-- 書籤 -->
    <w:bookmarkEnd/>
</w:p>
```

詳見：[11-paragraph.md](11-paragraph.md)

### w:tbl（表格）

表格結構。

```xml
<w:tbl>
    <w:tblPr>...</w:tblPr>     <!-- 表格屬性 -->
    <w:tblGrid>...</w:tblGrid> <!-- 欄寬定義 -->
    <w:tr>...</w:tr>           <!-- 表格列 -->
</w:tbl>
```

詳見：[20-table.md](20-table.md)

### w:sdt（結構化文件標籤）

Content Control，用於表單和範本。

```xml
<w:sdt>
    <w:sdtPr>
        <w:alias w:val="欄位名稱"/>
        <w:tag w:val="tag_name"/>
        <w:text/>  <!-- 或 w:richText, w:date, w:comboBox, etc. -->
    </w:sdtPr>
    <w:sdtContent>
        <w:p>
            <w:r>
                <w:t>預設值</w:t>
            </w:r>
        </w:p>
    </w:sdtContent>
</w:sdt>
```

---

## w:sectPr（分節屬性）

定義頁面設定和節的屬性。

### 位置

- **文件末節**：放在 `w:body` 的最後一個子元素
- **中間節**：放在段落的 `w:pPr` 中

### 子元素

| 元素 | 說明 |
|------|------|
| `w:pgSz` | 頁面大小 |
| `w:pgMar` | 頁邊距 |
| `w:pgBorders` | 頁面邊框 |
| `w:lnNumType` | 行號設定 |
| `w:pgNumType` | 頁碼設定 |
| `w:cols` | 欄設定 |
| `w:docGrid` | 文件格線 |
| `w:headerReference` | 頁首參照 |
| `w:footerReference` | 頁尾參照 |
| `w:type` | 分節類型 |

### 完整範例

```xml
<w:sectPr>
    <!-- 頁面大小 -->
    <w:pgSz w:w="12240" w:h="15840" w:orient="portrait" w:code="1"/>

    <!-- 頁邊距 -->
    <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"
             w:header="720" w:footer="720" w:gutter="0"/>

    <!-- 頁面邊框 -->
    <w:pgBorders w:offsetFrom="page">
        <w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/>
        <w:left w:val="single" w:sz="4" w:space="24" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:space="24" w:color="auto"/>
        <w:right w:val="single" w:sz="4" w:space="24" w:color="auto"/>
    </w:pgBorders>

    <!-- 行號 -->
    <w:lnNumType w:countBy="1" w:start="1" w:restart="newPage"/>

    <!-- 頁碼設定 -->
    <w:pgNumType w:fmt="decimal" w:start="1"/>

    <!-- 欄設定（兩欄） -->
    <w:cols w:num="2" w:space="720" w:equalWidth="1"/>

    <!-- 文件格線 -->
    <w:docGrid w:type="lines" w:linePitch="360"/>

    <!-- 頁首 -->
    <w:headerReference w:type="default" r:id="rId10"/>
    <w:headerReference w:type="first" r:id="rId11"/>
    <w:headerReference w:type="even" r:id="rId12"/>

    <!-- 頁尾 -->
    <w:footerReference w:type="default" r:id="rId13"/>
    <w:footerReference w:type="first" r:id="rId14"/>
    <w:footerReference w:type="even" r:id="rId15"/>

    <!-- 分節類型 -->
    <w:type w:val="nextPage"/>

    <!-- 首頁不同 -->
    <w:titlePg/>
</w:sectPr>
```

---

## 分節類型

| 值 | 說明 |
|----|------|
| `continuous` | 連續（同一頁開始新節） |
| `nextPage` | 下一頁 |
| `evenPage` | 偶數頁 |
| `oddPage` | 奇數頁 |
| `nextColumn` | 下一欄 |

---

## 下一步

- [11-paragraph.md](11-paragraph.md) - 段落元素詳解
- [20-table.md](20-table.md) - 表格結構
- [40-section.md](40-section.md) - 分節屬性詳解
