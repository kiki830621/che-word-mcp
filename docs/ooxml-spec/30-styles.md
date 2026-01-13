# 樣式 (Styles) 系統

## 概述

樣式系統定義在 `word/styles.xml` 中，提供可重用的格式設定。

## 基本結構

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

    <!-- 文件預設值 -->
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <!-- 預設文字屬性 -->
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault>
            <w:pPr>
                <!-- 預設段落屬性 -->
            </w:pPr>
        </w:pPrDefault>
    </w:docDefaults>

    <!-- 潛在樣式設定 -->
    <w:latentStyles>...</w:latentStyles>

    <!-- 樣式定義 -->
    <w:style w:type="paragraph" w:styleId="Normal">...</w:style>
    <w:style w:type="paragraph" w:styleId="Heading1">...</w:style>
    <w:style w:type="character" w:styleId="DefaultParagraphFont">...</w:style>
    <w:style w:type="table" w:styleId="TableNormal">...</w:style>
    <w:style w:type="numbering" w:styleId="NoList">...</w:style>

</w:styles>
```

---

## w:docDefaults（文件預設值）

定義整個文件的預設格式。

```xml
<w:docDefaults>
    <w:rPrDefault>
        <w:rPr>
            <w:rFonts w:ascii="Calibri" w:eastAsia="新細明體"
                      w:hAnsi="Calibri" w:cs="Times New Roman"/>
            <w:sz w:val="22"/>
            <w:szCs w:val="22"/>
            <w:lang w:val="en-US" w:eastAsia="zh-TW" w:bidi="ar-SA"/>
        </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
        <w:pPr>
            <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
        </w:pPr>
    </w:pPrDefault>
</w:docDefaults>
```

---

## w:style（樣式定義）

### 樣式屬性

| 屬性 | 說明 |
|------|------|
| `w:type` | 樣式類型 |
| `w:styleId` | 樣式 ID（用於參照） |
| `w:default` | 是否為預設樣式 |
| `w:customStyle` | 是否為自訂樣式 |

### 樣式類型 (w:type)

| 值 | 說明 | 應用於 |
|----|------|--------|
| `paragraph` | 段落樣式 | 段落和文字 |
| `character` | 字元樣式 | 僅文字 |
| `table` | 表格樣式 | 表格 |
| `numbering` | 編號樣式 | 編號/清單 |

### w:style 子元素

| 元素 | 說明 |
|------|------|
| `w:name` | 樣式名稱（顯示用） |
| `w:aliases` | 樣式別名 |
| `w:basedOn` | 基於（繼承） |
| `w:next` | 下一段落樣式 |
| `w:link` | 連結的樣式 |
| `w:autoRedefine` | 自動重新定義 |
| `w:hidden` | 隱藏樣式 |
| `w:uiPriority` | UI 優先順序 |
| `w:semiHidden` | 半隱藏 |
| `w:unhideWhenUsed` | 使用時取消隱藏 |
| `w:qFormat` | 快速樣式 |
| `w:locked` | 鎖定樣式 |
| `w:personal` | 個人樣式 |
| `w:personalCompose` | 撰寫時個人樣式 |
| `w:personalReply` | 回覆時個人樣式 |
| `w:rsid` | 修訂識別碼 |
| `w:pPr` | 段落屬性 |
| `w:rPr` | 文字屬性 |
| `w:tblPr` | 表格屬性 |
| `w:trPr` | 列屬性 |
| `w:tcPr` | 儲存格屬性 |
| `w:tblStylePr` | 條件式表格格式 |

---

## 段落樣式範例

### Normal（內文）

```xml
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
        <w:rFonts w:ascii="Calibri" w:eastAsia="新細明體" w:hAnsi="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
    </w:rPr>
</w:style>
```

### Heading1（標題 1）

```xml
<w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
        <w:keepNext/>
        <w:keepLines/>
        <w:spacing w:before="480" w:after="0"/>
        <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
        <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"
                  w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
        <w:b/>
        <w:bCs/>
        <w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
        <w:sz w:val="32"/>
        <w:szCs w:val="32"/>
    </w:rPr>
</w:style>
```

### Title（標題）

```xml
<w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="TitleChar"/>
    <w:uiPriority w:val="10"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
        <w:contextualSpacing/>
    </w:pPr>
    <w:rPr>
        <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"
                  w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
        <w:spacing w:val="-10"/>
        <w:kern w:val="28"/>
        <w:sz w:val="56"/>
        <w:szCs w:val="56"/>
    </w:rPr>
</w:style>
```

### Quote（引言）

```xml
<w:style w:type="paragraph" w:styleId="Quote">
    <w:name w:val="Quote"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="QuoteChar"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:before="200" w:after="160"/>
        <w:ind w:left="864" w:right="864"/>
        <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
        <w:i/>
        <w:iCs/>
        <w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/>
    </w:rPr>
</w:style>
```

### ListParagraph（清單段落）

```xml
<w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="34"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:left="720"/>
        <w:contextualSpacing/>
    </w:pPr>
</w:style>
```

---

## 字元樣式範例

### DefaultParagraphFont

```xml
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
</w:style>
```

### Hyperlink（超連結）

```xml
<w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="99"/>
    <w:unhideWhenUsed/>
    <w:rPr>
        <w:color w:val="0563C1" w:themeColor="hyperlink"/>
        <w:u w:val="single"/>
    </w:rPr>
</w:style>
```

### Strong（強調）

```xml
<w:style w:type="character" w:styleId="Strong">
    <w:name w:val="Strong"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="22"/>
    <w:qFormat/>
    <w:rPr>
        <w:b/>
        <w:bCs/>
    </w:rPr>
</w:style>
```

### Emphasis（斜體強調）

```xml
<w:style w:type="character" w:styleId="Emphasis">
    <w:name w:val="Emphasis"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="20"/>
    <w:qFormat/>
    <w:rPr>
        <w:i/>
        <w:iCs/>
    </w:rPr>
</w:style>
```

### 連結樣式

段落樣式可以有對應的字元樣式：

```xml
<!-- 段落樣式 -->
<w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:link w:val="Heading1Char"/>
    <!-- ... -->
</w:style>

<!-- 對應的字元樣式 -->
<w:style w:type="character" w:styleId="Heading1Char">
    <w:name w:val="Heading 1 Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="Heading1"/>
    <w:uiPriority w:val="9"/>
    <w:rPr>
        <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"/>
        <w:b/>
        <w:bCs/>
        <w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
        <w:sz w:val="32"/>
        <w:szCs w:val="32"/>
    </w:rPr>
</w:style>
```

---

## 表格樣式範例

### TableNormal

```xml
<w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
        <w:tblInd w:w="0" w:type="dxa"/>
        <w:tblCellMar>
            <w:top w:w="0" w:type="dxa"/>
            <w:left w:w="108" w:type="dxa"/>
            <w:bottom w:w="0" w:type="dxa"/>
            <w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar>
    </w:tblPr>
</w:style>
```

### TableGrid

```xml
<w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="39"/>
    <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:tblPr>
        <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>
    </w:tblPr>
</w:style>
```

### 帶條件格式的表格樣式

```xml
<w:style w:type="table" w:styleId="ColorfulShading">
    <w:name w:val="Colorful Shading"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="71"/>

    <!-- 基本表格屬性 -->
    <w:tblPr>
        <w:tblStyleRowBandSize w:val="1"/>
        <w:tblStyleColBandSize w:val="1"/>
        <w:tblBorders>
            <w:top w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>
            <w:bottom w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>
        </w:tblBorders>
    </w:tblPr>

    <!-- 條件式格式 -->
    <w:tblStylePr w:type="firstRow">
        <w:pPr>
            <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:bCs/>
            <w:color w:val="FFFFFF"/>
        </w:rPr>
        <w:tblPr/>
        <w:tcPr>
            <w:tcBorders>
                <w:top w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>
                <w:bottom w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>
            </w:tcBorders>
            <w:shd w:val="clear" w:color="auto" w:fill="4472C4"/>
        </w:tcPr>
    </w:tblStylePr>

    <w:tblStylePr w:type="lastRow">
        <w:rPr>
            <w:b/>
            <w:bCs/>
        </w:rPr>
        <w:tblPr/>
        <w:tcPr>
            <w:tcBorders>
                <w:top w:val="double" w:sz="6" w:space="0" w:color="4472C4"/>
            </w:tcBorders>
        </w:tcPr>
    </w:tblStylePr>

    <w:tblStylePr w:type="band1Horz">
        <w:tblPr/>
        <w:tcPr>
            <w:shd w:val="clear" w:color="auto" w:fill="D6DCE4"/>
        </w:tcPr>
    </w:tblStylePr>
</w:style>
```

### w:tblStylePr 條件類型

| w:type | 說明 |
|--------|------|
| `firstRow` | 首列 |
| `lastRow` | 末列 |
| `firstCol` | 首欄 |
| `lastCol` | 末欄 |
| `band1Vert` | 奇數垂直帶 |
| `band2Vert` | 偶數垂直帶 |
| `band1Horz` | 奇數水平帶 |
| `band2Horz` | 偶數水平帶 |
| `neCell` | 右上角儲存格 |
| `nwCell` | 左上角儲存格 |
| `seCell` | 右下角儲存格 |
| `swCell` | 左下角儲存格 |

---

## 編號樣式範例

```xml
<w:style w:type="numbering" w:default="1" w:styleId="NoList">
    <w:name w:val="No List"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
</w:style>
```

---

## w:latentStyles（潛在樣式）

控制內建樣式的預設行為。

```xml
<w:latentStyles w:defLockedState="0"
                w:defUIPriority="99"
                w:defSemiHidden="0"
                w:defUnhideWhenUsed="0"
                w:defQFormat="0"
                w:count="376">

    <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
    <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9"
                    w:unhideWhenUsed="1" w:qFormat="1"/>
    <w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>
    <w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>
    <w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>
    <w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>
    <!-- ... -->
</w:latentStyles>
```

### w:latentStyles 屬性

| 屬性 | 說明 |
|------|------|
| `w:defLockedState` | 預設鎖定狀態 |
| `w:defUIPriority` | 預設 UI 優先順序 |
| `w:defSemiHidden` | 預設半隱藏 |
| `w:defUnhideWhenUsed` | 預設使用時取消隱藏 |
| `w:defQFormat` | 預設快速樣式 |
| `w:count` | 內建樣式數量 |

### w:lsdException 屬性

| 屬性 | 說明 |
|------|------|
| `w:name` | 樣式名稱 |
| `w:locked` | 鎖定 |
| `w:uiPriority` | UI 優先順序 |
| `w:semiHidden` | 半隱藏 |
| `w:unhideWhenUsed` | 使用時取消隱藏 |
| `w:qFormat` | 快速樣式 |

---

## 樣式繼承

```
docDefaults
    ↓
Normal (段落樣式基礎)
    ↓
Heading1 (basedOn="Normal")
    ↓
自訂樣式 (basedOn="Heading1")
```

### 繼承範例

```xml
<!-- 基礎樣式 -->
<w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr>
        <w:spacing w:after="200"/>
    </w:pPr>
    <w:rPr>
        <w:sz w:val="22"/>
    </w:rPr>
</w:style>

<!-- 繼承並覆寫 -->
<w:style w:type="paragraph" w:styleId="CustomStyle">
    <w:name w:val="Custom Style"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
        <w:spacing w:after="0"/>  <!-- 覆寫段後間距 -->
        <w:ind w:left="720"/>     <!-- 新增縮排 -->
    </w:pPr>
    <!-- 字型大小繼承自 Normal -->
</w:style>
```

---

## 完整 styles.xml 範例

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">

    <!-- 文件預設值 -->
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"
                          w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
                <w:sz w:val="22"/>
                <w:szCs w:val="22"/>
                <w:lang w:val="en-US" w:eastAsia="zh-TW" w:bidi="ar-SA"/>
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault>
            <w:pPr>
                <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
            </w:pPr>
        </w:pPrDefault>
    </w:docDefaults>

    <!-- 潛在樣式 -->
    <w:latentStyles w:defLockedState="0" w:defUIPriority="99"
                    w:defSemiHidden="0" w:defUnhideWhenUsed="0"
                    w:defQFormat="0" w:count="376">
        <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
        <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
        <w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>
    </w:latentStyles>

    <!-- Normal -->
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:qFormat/>
    </w:style>

    <!-- Heading 1 -->
    <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading1Char"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
            <w:keepNext/>
            <w:keepLines/>
            <w:spacing w:before="240" w:after="0"/>
            <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"
                      w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
            <w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
            <w:sz w:val="32"/>
            <w:szCs w:val="32"/>
        </w:rPr>
    </w:style>

    <!-- Default Paragraph Font -->
    <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
        <w:name w:val="Default Paragraph Font"/>
        <w:uiPriority w:val="1"/>
        <w:semiHidden/>
        <w:unhideWhenUsed/>
    </w:style>

    <!-- Table Normal -->
    <w:style w:type="table" w:default="1" w:styleId="TableNormal">
        <w:name w:val="Normal Table"/>
        <w:uiPriority w:val="99"/>
        <w:semiHidden/>
        <w:unhideWhenUsed/>
        <w:tblPr>
            <w:tblInd w:w="0" w:type="dxa"/>
            <w:tblCellMar>
                <w:top w:w="0" w:type="dxa"/>
                <w:left w:w="108" w:type="dxa"/>
                <w:bottom w:w="0" w:type="dxa"/>
                <w:right w:w="108" w:type="dxa"/>
            </w:tblCellMar>
        </w:tblPr>
    </w:style>

    <!-- Table Grid -->
    <w:style w:type="table" w:styleId="TableGrid">
        <w:name w:val="Table Grid"/>
        <w:basedOn w:val="TableNormal"/>
        <w:uiPriority w:val="39"/>
        <w:pPr>
            <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
        </w:pPr>
        <w:tblPr>
            <w:tblBorders>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            </w:tblBorders>
        </w:tblPr>
    </w:style>

    <!-- No List -->
    <w:style w:type="numbering" w:default="1" w:styleId="NoList">
        <w:name w:val="No List"/>
        <w:uiPriority w:val="99"/>
        <w:semiHidden/>
        <w:unhideWhenUsed/>
    </w:style>

</w:styles>
```

---

## 下一步

- [31-style-types.md](31-style-types.md) - 各類型樣式詳解
- [32-numbering.md](32-numbering.md) - 編號定義
- [40-section.md](40-section.md) - 分節屬性
