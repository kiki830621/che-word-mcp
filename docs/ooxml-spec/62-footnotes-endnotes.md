# 腳註與尾註 (Footnotes & Endnotes)

## 概述

腳註（Footnotes）和尾註（Endnotes）是用於補充說明或引用來源的附加文字。腳註出現在頁面底部，尾註出現在文件或分節的末尾。

## 檔案結構

```
document.docx
├── word/
│   ├── document.xml      # 主文件（包含參照）
│   ├── footnotes.xml     # 腳註內容
│   ├── endnotes.xml      # 尾註內容
│   └── _rels/
│       └── document.xml.rels
└── [Content_Types].xml
```

## Content Types 定義

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/word/footnotes.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
    <Override PartName="/word/endnotes.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>
```

## 關聯定義

```xml
<!-- word/_rels/document.xml.rels -->
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
                  Target="footnotes.xml"/>
    <Relationship Id="rId4"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
                  Target="endnotes.xml"/>
</Relationships>
```

---

## 腳註 (Footnotes)

### footnotes.xml 基本結構

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

    <!-- 分隔符腳註（系統使用） -->
    <w:footnote w:type="separator" w:id="-1">
        <w:p>
            <w:r>
                <w:separator/>
            </w:r>
        </w:p>
    </w:footnote>

    <!-- 延續分隔符（跨頁時使用） -->
    <w:footnote w:type="continuationSeparator" w:id="0">
        <w:p>
            <w:r>
                <w:continuationSeparator/>
            </w:r>
        </w:p>
    </w:footnote>

    <!-- 實際腳註內容 -->
    <w:footnote w:id="1">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="FootnoteText"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="FootnoteReference"/>
                </w:rPr>
                <w:footnoteRef/>
            </w:r>
            <w:r>
                <w:t xml:space="preserve"> </w:t>
            </w:r>
            <w:r>
                <w:t>這是腳註的內容說明。</w:t>
            </w:r>
        </w:p>
    </w:footnote>

</w:footnotes>
```

### w:footnote 屬性

| 屬性 | 說明 | 值 |
|------|------|-----|
| `w:id` | 腳註識別碼 | 整數（-1, 0 為系統保留） |
| `w:type` | 腳註類型 | `separator`, `continuationSeparator`, `continuationNotice` |

### 腳註類型

| 類型 | 說明 |
|------|------|
| `separator` | 正文與腳註之間的分隔線 |
| `continuationSeparator` | 腳註延續到下頁時的分隔線 |
| `continuationNotice` | 延續提示文字 |
| （無 type） | 一般腳註內容 |

### 在 document.xml 中參照腳註

```xml
<w:p>
    <w:r>
        <w:t>這是正文內容</w:t>
    </w:r>
    <w:r>
        <w:rPr>
            <w:rStyle w:val="FootnoteReference"/>
        </w:rPr>
        <w:footnoteReference w:id="1"/>
    </w:r>
    <w:r>
        <w:t>，腳註已加入。</w:t>
    </w:r>
</w:p>
```

---

## 尾註 (Endnotes)

### endnotes.xml 基本結構

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

    <!-- 分隔符（系統使用） -->
    <w:endnote w:type="separator" w:id="-1">
        <w:p>
            <w:r>
                <w:separator/>
            </w:r>
        </w:p>
    </w:endnote>

    <w:endnote w:type="continuationSeparator" w:id="0">
        <w:p>
            <w:r>
                <w:continuationSeparator/>
            </w:r>
        </w:p>
    </w:endnote>

    <!-- 實際尾註內容 -->
    <w:endnote w:id="1">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="EndnoteText"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="EndnoteReference"/>
                </w:rPr>
                <w:endnoteRef/>
            </w:r>
            <w:r>
                <w:t xml:space="preserve"> </w:t>
            </w:r>
            <w:r>
                <w:t>參見《資料來源》第 42 頁。</w:t>
            </w:r>
        </w:p>
    </w:endnote>

</w:endnotes>
```

### 在 document.xml 中參照尾註

```xml
<w:p>
    <w:r>
        <w:t>根據研究結果</w:t>
    </w:r>
    <w:r>
        <w:rPr>
            <w:rStyle w:val="EndnoteReference"/>
        </w:rPr>
        <w:endnoteReference w:id="1"/>
    </w:r>
    <w:r>
        <w:t>顯示...</w:t>
    </w:r>
</w:p>
```

---

## 腳註/尾註設定

### 在 settings.xml 中設定

```xml
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

    <!-- 腳註設定 -->
    <w:footnotePr>
        <!-- 位置：pageBottom（頁底）或 beneathText（文字下方） -->
        <w:pos w:val="pageBottom"/>

        <!-- 編號格式 -->
        <w:numFmt w:val="decimal"/>

        <!-- 起始編號 -->
        <w:numStart w:val="1"/>

        <!-- 重新編號：continuous（連續）、eachSect（每節）、eachPage（每頁） -->
        <w:numRestart w:val="continuous"/>
    </w:footnotePr>

    <!-- 尾註設定 -->
    <w:endnotePr>
        <!-- 位置：docEnd（文件末）或 sectEnd（節末） -->
        <w:pos w:val="docEnd"/>

        <!-- 編號格式 -->
        <w:numFmt w:val="lowerRoman"/>

        <!-- 起始編號 -->
        <w:numStart w:val="1"/>

        <!-- 重新編號 -->
        <w:numRestart w:val="eachSect"/>
    </w:endnotePr>

</w:settings>
```

### 在分節屬性中覆寫設定

```xml
<w:sectPr>
    <!-- 覆寫此分節的腳註設定 -->
    <w:footnotePr>
        <w:numFmt w:val="upperLetter"/>
        <w:numStart w:val="1"/>
        <w:numRestart w:val="eachSect"/>
    </w:footnotePr>

    <!-- 覆寫此分節的尾註設定 -->
    <w:endnotePr>
        <w:pos w:val="sectEnd"/>  <!-- 在本節末顯示 -->
    </w:endnotePr>

    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>
</w:sectPr>
```

### 編號格式 (w:numFmt)

| 值 | 說明 | 範例 |
|-----|------|------|
| `decimal` | 阿拉伯數字 | 1, 2, 3 |
| `upperRoman` | 大寫羅馬數字 | I, II, III |
| `lowerRoman` | 小寫羅馬數字 | i, ii, iii |
| `upperLetter` | 大寫字母 | A, B, C |
| `lowerLetter` | 小寫字母 | a, b, c |
| `chicago` | 芝加哥格式符號 | *, †, ‡ |

### 腳註位置 (footnotePr/pos)

| 值 | 說明 |
|-----|------|
| `pageBottom` | 頁面底部（預設） |
| `beneathText` | 緊接在內文下方 |

### 尾註位置 (endnotePr/pos)

| 值 | 說明 |
|-----|------|
| `docEnd` | 文件末尾（預設） |
| `sectEnd` | 分節末尾 |

---

## 樣式定義

### styles.xml 中的預設樣式

```xml
<!-- 腳註文字樣式 -->
<w:style w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="FootnoteTextChar"/>
    <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
    </w:rPr>
</w:style>

<!-- 腳註參照樣式 -->
<w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
    <w:rPr>
        <w:vertAlign w:val="superscript"/>
    </w:rPr>
</w:style>

<!-- 尾註文字樣式 -->
<w:style w:type="paragraph" w:styleId="EndnoteText">
    <w:name w:val="endnote text"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="EndnoteTextChar"/>
    <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
    </w:rPr>
</w:style>

<!-- 尾註參照樣式 -->
<w:style w:type="character" w:styleId="EndnoteReference">
    <w:name w:val="endnote reference"/>
    <w:rPr>
        <w:vertAlign w:val="superscript"/>
    </w:rPr>
</w:style>
```

---

## 完整範例

### footnotes.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

    <w:footnote w:type="separator" w:id="-1">
        <w:p>
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
            <w:r>
                <w:separator/>
            </w:r>
        </w:p>
    </w:footnote>

    <w:footnote w:type="continuationSeparator" w:id="0">
        <w:p>
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
            <w:r>
                <w:continuationSeparator/>
            </w:r>
        </w:p>
    </w:footnote>

    <!-- 第一個腳註 -->
    <w:footnote w:id="1">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="FootnoteText"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="FootnoteReference"/>
                </w:rPr>
                <w:footnoteRef/>
            </w:r>
            <w:r>
                <w:t xml:space="preserve"> </w:t>
            </w:r>
            <w:r>
                <w:t>此數據來自 2023 年全國調查報告。</w:t>
            </w:r>
        </w:p>
    </w:footnote>

    <!-- 第二個腳註（多段落） -->
    <w:footnote w:id="2">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="FootnoteText"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="FootnoteReference"/>
                </w:rPr>
                <w:footnoteRef/>
            </w:r>
            <w:r>
                <w:t xml:space="preserve"> </w:t>
            </w:r>
            <w:r>
                <w:t>關於此議題的詳細討論，請參考：</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="FootnoteText"/>
                <w:ind w:left="720"/>
            </w:pPr>
            <w:r>
                <w:t>張三，《研究方法論》，2022年，第 45-67 頁。</w:t>
            </w:r>
        </w:p>
    </w:footnote>

</w:footnotes>
```

### document.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:r>
                <w:t>研究背景</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>根據最新統計資料</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="FootnoteReference"/>
                </w:rPr>
                <w:footnoteReference w:id="1"/>
            </w:r>
            <w:r>
                <w:t>，有超過 80% 的受訪者表示...</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>這個觀點在學術界已有廣泛討論</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="FootnoteReference"/>
                </w:rPr>
                <w:footnoteReference w:id="2"/>
            </w:r>
            <w:r>
                <w:t>。</w:t>
            </w:r>
        </w:p>

        <w:sectPr>
            <w:footnotePr>
                <w:numFmt w:val="decimal"/>
            </w:footnotePr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>
        </w:sectPr>
    </w:body>
</w:document>
```

---

## 自訂分隔符

### 自訂腳註分隔線

```xml
<w:footnote w:type="separator" w:id="-1">
    <w:p>
        <w:pPr>
            <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
        </w:pPr>
        <w:r>
            <w:rPr>
                <w:sz w:val="16"/>
            </w:rPr>
            <!-- 使用水平線字元 -->
            <w:t>────────────</w:t>
        </w:r>
    </w:p>
</w:footnote>
```

### 自訂延續提示

```xml
<w:footnote w:type="continuationNotice" w:id="1">
    <w:p>
        <w:pPr>
            <w:jc w:val="right"/>
        </w:pPr>
        <w:r>
            <w:rPr>
                <w:i/>
                <w:sz w:val="18"/>
            </w:rPr>
            <w:t>（續下頁）</w:t>
        </w:r>
    </w:p>
</w:footnote>
```

---

## 實作注意事項

### ID 規則
- ID `-1` 和 `0` 是系統保留的（分隔符用）
- 一般腳註/尾註 ID 從 `1` 開始
- ID 必須在各自的檔案中唯一（footnotes 和 endnotes 分開計算）

### 必要的分隔符
- `separator`（ID=-1）和 `continuationSeparator`（ID=0）是必要的
- 即使文件沒有腳註/尾註，這兩個仍應存在於 XML 中

### 參照元素
- 文件中使用 `w:footnoteReference` / `w:endnoteReference`
- 腳註/尾註內容中使用 `w:footnoteRef` / `w:endnoteRef`（顯示編號）

### 樣式
- 建議定義 FootnoteText 和 EndnoteText 段落樣式
- 建議定義 FootnoteReference 和 EndnoteReference 字元樣式

---

## 相關連結

- [段落結構](11-paragraph.md)
- [樣式系統](30-styles.md)
- [分節屬性](40-section.md)
- [欄位代碼](64-fields.md)
