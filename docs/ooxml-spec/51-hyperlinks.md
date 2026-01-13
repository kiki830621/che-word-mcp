# 超連結與書籤 (Hyperlinks & Bookmarks)

## 概述

OOXML 支援兩種連結類型：
1. **超連結**（Hyperlink）- 連結到外部 URL 或文件內部位置
2. **書籤**（Bookmark）- 文件內的標記點，可被參照

---

## 超連結 (Hyperlink)

### 外部超連結

#### 關聯定義

```xml
<!-- word/_rels/document.xml.rels -->
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId5"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                  Target="https://www.example.com"
                  TargetMode="External"/>
</Relationships>
```

#### 在 document.xml 中使用

```xml
<w:p>
    <w:hyperlink r:id="rId5" w:history="1">
        <w:r>
            <w:rPr>
                <w:rStyle w:val="Hyperlink"/>
            </w:rPr>
            <w:t>點擊這裡前往網站</w:t>
        </w:r>
    </w:hyperlink>
</w:p>
```

### 內部超連結（連結到書籤）

```xml
<w:p>
    <w:hyperlink w:anchor="Chapter1">
        <w:r>
            <w:rPr>
                <w:rStyle w:val="Hyperlink"/>
            </w:rPr>
            <w:t>前往第一章</w:t>
        </w:r>
    </w:hyperlink>
</w:p>
```

### w:hyperlink 屬性

| 屬性 | 說明 |
|------|------|
| `r:id` | 關聯 ID（外部連結） |
| `w:anchor` | 書籤名稱（內部連結） |
| `w:history` | 是否加入歷史記錄 |
| `w:tooltip` | 滑鼠提示文字 |
| `w:docLocation` | 文件位置 |
| `w:tgtFrame` | 目標框架 |

### 含提示文字的超連結

```xml
<w:hyperlink r:id="rId5" w:tooltip="這是提示文字" w:history="1">
    <w:r>
        <w:rPr>
            <w:rStyle w:val="Hyperlink"/>
        </w:rPr>
        <w:t>帶提示的連結</w:t>
    </w:r>
</w:hyperlink>
```

### 使用欄位代碼的超連結

```xml
<w:p>
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText xml:space="preserve"> HYPERLINK "https://www.example.com" \o "點擊前往" </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
        <w:rPr>
            <w:rStyle w:val="Hyperlink"/>
        </w:rPr>
        <w:t>連結文字</w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
</w:p>
```

---

## 超連結樣式

### 預設樣式定義

```xml
<!-- styles.xml -->
<w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:rPr>
        <w:color w:val="0563C1" w:themeColor="hyperlink"/>
        <w:u w:val="single"/>
    </w:rPr>
</w:style>

<w:style w:type="character" w:styleId="FollowedHyperlink">
    <w:name w:val="FollowedHyperlink"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:rPr>
        <w:color w:val="954F72" w:themeColor="followedHyperlink"/>
        <w:u w:val="single"/>
    </w:rPr>
</w:style>
```

---

## 書籤 (Bookmark)

### 基本書籤

```xml
<w:p>
    <w:bookmarkStart w:id="0" w:name="Introduction"/>
    <w:r>
        <w:t>這是書籤標記的內容</w:t>
    </w:r>
    <w:bookmarkEnd w:id="0"/>
</w:p>
```

### 書籤屬性

| 元素 | 屬性 | 說明 |
|------|------|------|
| `w:bookmarkStart` | `w:id` | 書籤 ID（與 bookmarkEnd 配對） |
| | `w:name` | 書籤名稱 |
| | `w:colFirst` | 表格列起始（表格書籤用） |
| | `w:colLast` | 表格列結束（表格書籤用） |
| `w:bookmarkEnd` | `w:id` | 對應的書籤 ID |

### 跨段落書籤

```xml
<w:p>
    <w:bookmarkStart w:id="0" w:name="Section1"/>
    <w:r>
        <w:t>第一段</w:t>
    </w:r>
</w:p>
<w:p>
    <w:r>
        <w:t>第二段</w:t>
    </w:r>
</w:p>
<w:p>
    <w:r>
        <w:t>第三段</w:t>
    </w:r>
    <w:bookmarkEnd w:id="0"/>
</w:p>
```

### 空書籤（位置標記）

```xml
<w:p>
    <w:r>
        <w:t>文字</w:t>
    </w:r>
    <w:bookmarkStart w:id="0" w:name="InsertPoint"/>
    <w:bookmarkEnd w:id="0"/>
    <w:r>
        <w:t>更多文字</w:t>
    </w:r>
</w:p>
```

### 隱藏書籤

以底線開頭的書籤名稱會被隱藏：

```xml
<w:bookmarkStart w:id="0" w:name="_Toc123456789"/>
```

常見隱藏書籤：
- `_Toc*` - 目錄書籤
- `_Ref*` - 交互參照
- `_GoBack` - 上次編輯位置

---

## 表格書籤

### 標記表格儲存格範圍

```xml
<w:tbl>
    <w:tr>
        <w:tc>
            <w:p>
                <w:bookmarkStart w:id="0" w:name="TableData"
                                 w:colFirst="1" w:colLast="3"/>
                <w:r><w:t>A1</w:t></w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:p>
                <w:r><w:t>B1</w:t></w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:p>
                <w:r><w:t>C1</w:t></w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:p>
                <w:r><w:t>D1</w:t></w:r>
                <w:bookmarkEnd w:id="0"/>
            </w:p>
        </w:tc>
    </w:tr>
</w:tbl>
```

---

## 參照書籤

### REF 欄位（取得書籤內容）

```xml
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> REF Introduction </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>這是書籤標記的內容</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### PAGEREF 欄位（取得書籤頁碼）

```xml
<w:r>
    <w:t>詳見第 </w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> PAGEREF Introduction </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>5</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
<w:r>
    <w:t> 頁</w:t>
</w:r>
```

### REF 欄位開關

| 開關 | 說明 |
|------|------|
| `\f` | 包含腳註/尾註編號 |
| `\h` | 建立超連結 |
| `\n` | 插入段落編號 |
| `\p` | 插入相對位置（above/below） |
| `\r` | 插入完整段落編號 |
| `\t` | 抑制非分隔符字元 |
| `\w` | 插入完整段落編號（含內容） |

### PAGEREF 欄位開關

| 開關 | 說明 |
|------|------|
| `\h` | 建立超連結 |
| `\p` | 插入相對位置（on page X, above, below） |

---

## 完整範例

### document.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <!-- 目錄連結 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="TOC1"/>
            </w:pPr>
            <w:hyperlink w:anchor="Chapter1">
                <w:r>
                    <w:t>第一章 簡介</w:t>
                </w:r>
                <w:r>
                    <w:tab/>
                </w:r>
                <w:r>
                    <w:fldChar w:fldCharType="begin"/>
                </w:r>
                <w:r>
                    <w:instrText xml:space="preserve"> PAGEREF Chapter1 \h </w:instrText>
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
            </w:hyperlink>
        </w:p>

        <!-- 分頁 -->
        <w:p>
            <w:r>
                <w:br w:type="page"/>
            </w:r>
        </w:p>

        <!-- 第一章（含書籤） -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:bookmarkStart w:id="0" w:name="Chapter1"/>
            <w:r>
                <w:t>第一章 簡介</w:t>
            </w:r>
            <w:bookmarkEnd w:id="0"/>
        </w:p>

        <w:p>
            <w:r>
                <w:t>這是第一章的內容。如需了解更多資訊，請參考 </w:t>
            </w:r>
            <!-- 外部超連結 -->
            <w:hyperlink r:id="rId5" w:tooltip="前往官方網站">
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="Hyperlink"/>
                    </w:rPr>
                    <w:t>官方網站</w:t>
                </w:r>
            </w:hyperlink>
            <w:r>
                <w:t>。</w:t>
            </w:r>
        </w:p>

        <!-- 重要內容（書籤） -->
        <w:p>
            <w:bookmarkStart w:id="1" w:name="ImportantNote"/>
            <w:r>
                <w:rPr>
                    <w:b/>
                </w:rPr>
                <w:t>重要提示：請務必閱讀此內容。</w:t>
            </w:r>
            <w:bookmarkEnd w:id="1"/>
        </w:p>

        <!-- 交互參照 -->
        <w:p>
            <w:r>
                <w:t>如前所述（參見「</w:t>
            </w:r>
            <w:r>
                <w:fldChar w:fldCharType="begin"/>
            </w:r>
            <w:r>
                <w:instrText xml:space="preserve"> REF ImportantNote \h </w:instrText>
            </w:r>
            <w:r>
                <w:fldChar w:fldCharType="separate"/>
            </w:r>
            <w:hyperlink w:anchor="ImportantNote">
                <w:r>
                    <w:t>重要提示：請務必閱讀此內容。</w:t>
                </w:r>
            </w:hyperlink>
            <w:r>
                <w:fldChar w:fldCharType="end"/>
            </w:r>
            <w:r>
                <w:t>」），這是非常重要的。</w:t>
            </w:r>
        </w:p>

        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>
        </w:sectPr>
    </w:body>
</w:document>
```

### document.xml.rels

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                  Target="styles.xml"/>
    <Relationship Id="rId5"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                  Target="https://www.example.com"
                  TargetMode="External"/>
</Relationships>
```

---

## 電子郵件連結

```xml
<!-- document.xml.rels -->
<Relationship Id="rId6"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="mailto:support@example.com?subject=詢問"
              TargetMode="External"/>

<!-- document.xml -->
<w:hyperlink r:id="rId6">
    <w:r>
        <w:rPr>
            <w:rStyle w:val="Hyperlink"/>
        </w:rPr>
        <w:t>聯絡我們</w:t>
    </w:r>
</w:hyperlink>
```

---

## 檔案連結

```xml
<!-- 連結到本機檔案 -->
<Relationship Id="rId7"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="file:///C:/Documents/report.pdf"
              TargetMode="External"/>

<!-- 連結到相對路徑 -->
<Relationship Id="rId8"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="../附件/資料.xlsx"
              TargetMode="External"/>
```

---

## 實作注意事項

### 書籤命名
- 名稱必須唯一
- 不能以數字開頭
- 避免使用特殊字元
- 以底線開頭會被隱藏

### ID 管理
- 書籤 ID 必須唯一
- 建議從 0 開始遞增
- bookmarkStart 和 bookmarkEnd 的 ID 必須配對

### 超連結關聯
- 外部連結需要在 .rels 中定義
- 必須設定 `TargetMode="External"`
- 內部連結使用 `w:anchor` 不需要關聯

### 相容性
- 確保書籤和超連結的目標存在
- 外部連結可能因網路或權限問題無法開啟

---

## 相關連結

- [欄位代碼](64-fields.md)
- [目錄](63-toc.md)
- [關聯](02-namespaces.md)
- [樣式系統](30-styles.md)
