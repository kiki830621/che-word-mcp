# 註解 (Comments)

## 概述

OOXML 支援在文件中加入註解（批註），允許多位使用者對文件內容進行評論和討論。註解儲存在獨立的 `comments.xml` 檔案中。

## 檔案結構

```
document.docx
├── word/
│   ├── document.xml      # 主文件（包含註解標記）
│   ├── comments.xml      # 註解內容
│   └── _rels/
│       └── document.xml.rels  # 關聯定義
└── [Content_Types].xml   # 內容類型
```

## Content Types 定義

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/word/comments.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>
```

## 關聯定義

```xml
<!-- word/_rels/document.xml.rels -->
<Relationship Id="rId5"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
              Target="comments.xml"/>
```

---

## comments.xml 結構

### 基本結構

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:comment w:id="0" w:author="張三" w:date="2024-01-15T10:30:00Z" w:initials="ZS">
        <w:p>
            <w:r>
                <w:t>這裡需要修改</w:t>
            </w:r>
        </w:p>
    </w:comment>
    <w:comment w:id="1" w:author="李四" w:date="2024-01-15T11:00:00Z" w:initials="LS">
        <w:p>
            <w:r>
                <w:t>同意，建議改成這樣...</w:t>
            </w:r>
        </w:p>
    </w:comment>
</w:comments>
```

### w:comment 屬性

| 屬性 | 說明 | 範例 |
|------|------|------|
| `w:id` | 註解唯一識別碼 | `"0"`, `"1"` |
| `w:author` | 作者名稱 | `"張三"` |
| `w:date` | 建立日期時間（ISO 8601） | `"2024-01-15T10:30:00Z"` |
| `w:initials` | 作者縮寫 | `"ZS"` |

---

## 在 document.xml 中標記註解範圍

### 註解範圍元素

| 元素 | 說明 |
|------|------|
| `w:commentRangeStart` | 註解範圍開始 |
| `w:commentRangeEnd` | 註解範圍結束 |
| `w:commentReference` | 註解參照（顯示註解編號） |

### 範例：標記單一段落

```xml
<w:p>
    <w:commentRangeStart w:id="0"/>
    <w:r>
        <w:t>這段文字有註解</w:t>
    </w:r>
    <w:commentRangeEnd w:id="0"/>
    <w:r>
        <w:commentReference w:id="0"/>
    </w:r>
</w:p>
```

### 範例：跨段落註解

```xml
<w:p>
    <w:commentRangeStart w:id="0"/>
    <w:r>
        <w:t>第一段開始...</w:t>
    </w:r>
</w:p>
<w:p>
    <w:r>
        <w:t>第二段內容...</w:t>
    </w:r>
    <w:commentRangeEnd w:id="0"/>
    <w:r>
        <w:commentReference w:id="0"/>
    </w:r>
</w:p>
```

### 範例：巢狀註解

```xml
<w:p>
    <w:commentRangeStart w:id="0"/>
    <w:r>
        <w:t>外層註解開始</w:t>
    </w:r>
    <w:commentRangeStart w:id="1"/>
    <w:r>
        <w:t>這裡有兩個註解</w:t>
    </w:r>
    <w:commentRangeEnd w:id="1"/>
    <w:r>
        <w:commentReference w:id="1"/>
    </w:r>
    <w:r>
        <w:t>外層註解結束</w:t>
    </w:r>
    <w:commentRangeEnd w:id="0"/>
    <w:r>
        <w:commentReference w:id="0"/>
    </w:r>
</w:p>
```

---

## 註解回覆 (Comment Replies)

### 擴展關聯

使用 Word 2013+ 格式，需要額外的擴展檔案：

```
document.docx
├── word/
│   ├── comments.xml              # 主要註解
│   ├── commentsExtended.xml      # 擴展資訊（回覆關係）
│   └── commentsIds.xml           # 註解 ID 對應
```

### commentsExtended.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <!-- 原始註解（無父級） -->
    <w15:commentEx w15:paraId="00000001" w15:done="0"/>

    <!-- 回覆註解（指向父級） -->
    <w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001" w15:done="0"/>
</w15:commentsEx>
```

### w15:commentEx 屬性

| 屬性 | 說明 |
|------|------|
| `w15:paraId` | 段落 ID（對應註解內的段落） |
| `w15:paraIdParent` | 父級註解的段落 ID |
| `w15:done` | 是否已解決（0=未解決，1=已解決） |

---

## 完整範例

### comments.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">

    <!-- 原始註解 -->
    <w:comment w:id="0" w:author="審稿人" w:date="2024-01-15T09:00:00Z" w:initials="S">
        <w:p w14:paraId="00000001">
            <w:pPr>
                <w:pStyle w:val="CommentText"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="CommentReference"/>
                </w:rPr>
                <w:annotationRef/>
            </w:r>
            <w:r>
                <w:t>這個數據需要確認來源</w:t>
            </w:r>
        </w:p>
    </w:comment>

    <!-- 回覆註解 -->
    <w:comment w:id="1" w:author="作者" w:date="2024-01-15T10:30:00Z" w:initials="A">
        <w:p w14:paraId="00000002">
            <w:pPr>
                <w:pStyle w:val="CommentText"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="CommentReference"/>
                </w:rPr>
                <w:annotationRef/>
            </w:r>
            <w:r>
                <w:t>已補充資料來源，請參考參考文獻第 3 項</w:t>
            </w:r>
        </w:p>
    </w:comment>

</w:comments>
```

### document.xml 中的標記

```xml
<w:body>
    <w:p>
        <w:r>
            <w:t>根據研究顯示，</w:t>
        </w:r>
        <w:commentRangeStart w:id="0"/>
        <w:r>
            <w:t>有 85% 的使用者表示滿意</w:t>
        </w:r>
        <w:commentRangeEnd w:id="0"/>
        <w:r>
            <w:commentReference w:id="0"/>
        </w:r>
        <w:r>
            <w:t>。</w:t>
        </w:r>
    </w:p>
</w:body>
```

---

## 註解樣式

### 預設樣式定義

```xml
<!-- styles.xml -->
<w:style w:type="paragraph" w:styleId="CommentText">
    <w:name w:val="annotation text"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="CommentTextChar"/>
    <w:rPr>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
    </w:rPr>
</w:style>

<w:style w:type="character" w:styleId="CommentReference">
    <w:name w:val="annotation reference"/>
    <w:rPr>
        <w:sz w:val="16"/>
        <w:szCs w:val="16"/>
    </w:rPr>
</w:style>

<w:style w:type="paragraph" w:styleId="CommentSubject">
    <w:name w:val="annotation subject"/>
    <w:basedOn w:val="CommentText"/>
    <w:next w:val="CommentText"/>
    <w:rPr>
        <w:b/>
        <w:bCs/>
    </w:rPr>
</w:style>
```

---

## 實作注意事項

### ID 管理
- 註解 ID 必須在整個文件中唯一
- ID 為非負整數，通常從 0 開始遞增
- 刪除註解後，ID 不需要重新編號

### 範圍匹配
- `commentRangeStart` 和 `commentRangeEnd` 的 `w:id` 必須匹配
- `commentReference` 的 `w:id` 也必須對應同一個註解

### 日期格式
- 使用 ISO 8601 格式
- 建議使用 UTC 時區（以 Z 結尾）
- 範例：`2024-01-15T10:30:00Z`

### 相容性
- 基本註解（comments.xml）被所有 Word 版本支援
- 回覆功能需要 Word 2013+ 的擴展格式
- 舊版 Word 會忽略擴展資訊，但仍顯示所有註解

---

## 相關連結

- [追蹤修訂](61-track-changes.md)
- [段落結構](11-paragraph.md)
- [樣式系統](30-styles.md)
