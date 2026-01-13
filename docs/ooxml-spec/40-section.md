# 分節屬性 (Section Properties)

## 概述

`w:sectPr` 定義文件節的頁面設定，包括頁面大小、邊距、方向、欄設定等。

## 位置

分節屬性可以出現在兩個位置：

1. **文件末節**：作為 `w:body` 的最後一個子元素
2. **中間節**：在段落的 `w:pPr` 中

```xml
<!-- 文件末節 -->
<w:body>
    <w:p>...</w:p>
    <w:sectPr>...</w:sectPr>
</w:body>

<!-- 中間節（在段落中） -->
<w:p>
    <w:pPr>
        <w:sectPr>...</w:sectPr>
    </w:pPr>
    <w:r>...</w:r>
</w:p>
```

---

## w:sectPr 子元素

| 元素 | 說明 |
|------|------|
| `w:headerReference` | 頁首參照 |
| `w:footerReference` | 頁尾參照 |
| `w:footnotePr` | 腳註屬性 |
| `w:endnotePr` | 尾註屬性 |
| `w:type` | 分節類型 |
| `w:pgSz` | 頁面大小 |
| `w:pgMar` | 頁邊距 |
| `w:paperSrc` | 紙張來源 |
| `w:pgBorders` | 頁面邊框 |
| `w:lnNumType` | 行號設定 |
| `w:pgNumType` | 頁碼設定 |
| `w:cols` | 欄設定 |
| `w:formProt` | 表單保護 |
| `w:vAlign` | 垂直對齊 |
| `w:noEndnote` | 無尾註 |
| `w:titlePg` | 首頁不同 |
| `w:textDirection` | 文字方向 |
| `w:bidi` | 雙向 |
| `w:rtlGutter` | 右到左裝訂 |
| `w:docGrid` | 文件格線 |
| `w:printerSettings` | 印表機設定 |

---

## w:pgSz（頁面大小）

```xml
<w:pgSz w:w="11906" w:h="16838" w:orient="portrait" w:code="9"/>
```

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:w` | 寬度（twips） |
| `w:h` | 高度（twips） |
| `w:orient` | 方向：`portrait`（直）、`landscape`（橫） |
| `w:code` | 紙張代碼 |

### 常用紙張大小

| 紙張 | 寬度 (twips) | 高度 (twips) | w:code |
|------|-------------|-------------|--------|
| Letter (8.5×11") | 12240 | 15840 | 1 |
| Legal (8.5×14") | 12240 | 20160 | 5 |
| A3 (297×420mm) | 16838 | 23811 | 8 |
| A4 (210×297mm) | 11906 | 16838 | 9 |
| A5 (148×210mm) | 8391 | 11906 | 11 |
| B4 (250×353mm) | 14173 | 20016 | 12 |
| B5 (176×250mm) | 9979 | 14173 | 13 |

### 換算公式

```
twips = inches × 1440
twips = cm × 567
twips = mm × 56.7
```

### 橫向頁面

```xml
<!-- 橫向 A4 -->
<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
```

**注意：** 橫向時，`w:w` 和 `w:h` 的值互換。

---

## w:pgMar（頁邊距）

```xml
<w:pgMar w:top="1440"
         w:right="1440"
         w:bottom="1440"
         w:left="1440"
         w:header="720"
         w:footer="720"
         w:gutter="0"/>
```

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:top` | 上邊距 |
| `w:bottom` | 下邊距 |
| `w:left` | 左邊距 |
| `w:right` | 右邊距 |
| `w:header` | 頁首距離（從頁面頂端） |
| `w:footer` | 頁尾距離（從頁面底端） |
| `w:gutter` | 裝訂邊 |

### 常用邊距設定

| 設定 | 上/下 | 左/右 | 說明 |
|------|-------|-------|------|
| 標準 | 1440 | 1440 | 1 inch |
| 窄 | 720 | 720 | 0.5 inch |
| 適中 | 1440 | 1080 | 1"/0.75" |
| 寬 | 1440 | 2880 | 1"/2" |

---

## w:type（分節類型）

```xml
<w:type w:val="nextPage"/>
```

| 值 | 說明 |
|----|------|
| `continuous` | 連續（同一頁開始新節） |
| `nextPage` | 下一頁（預設） |
| `evenPage` | 偶數頁 |
| `oddPage` | 奇數頁 |
| `nextColumn` | 下一欄 |

---

## w:cols（欄設定）

### 等寬欄

```xml
<!-- 兩欄等寬 -->
<w:cols w:num="2" w:space="720"/>

<!-- 三欄等寬 -->
<w:cols w:num="3" w:space="720"/>
```

### 不等寬欄

```xml
<w:cols w:num="2" w:space="720" w:equalWidth="0">
    <w:col w:w="5760" w:space="720"/>
    <w:col w:w="2880"/>
</w:cols>
```

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:num` | 欄數 |
| `w:space` | 欄間距（twips） |
| `w:equalWidth` | 等寬欄（0=否，1=是） |
| `w:sep` | 顯示分隔線 |

### w:col（個別欄設定）

| 屬性 | 說明 |
|------|------|
| `w:w` | 欄寬 |
| `w:space` | 此欄後的間距 |

---

## w:pgBorders（頁面邊框）

```xml
<w:pgBorders w:offsetFrom="page" w:display="allPages">
    <w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/>
    <w:left w:val="single" w:sz="4" w:space="24" w:color="auto"/>
    <w:bottom w:val="single" w:sz="4" w:space="24" w:color="auto"/>
    <w:right w:val="single" w:sz="4" w:space="24" w:color="auto"/>
</w:pgBorders>
```

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:offsetFrom` | 偏移起點：`page`（頁面）、`text`（文字） |
| `w:display` | 顯示頁面：`allPages`、`firstPage`、`notFirstPage` |
| `w:zOrder` | Z 順序：`front`、`back` |

### 邊框元素

每個邊框（top, left, bottom, right）支援：

| 屬性 | 說明 |
|------|------|
| `w:val` | 邊框樣式（參見段落邊框） |
| `w:sz` | 寬度（1/8 pt） |
| `w:space` | 間距（pt） |
| `w:color` | 顏色 |

---

## w:lnNumType（行號）

```xml
<w:lnNumType w:countBy="1"
             w:start="1"
             w:restart="newPage"
             w:distance="720"/>
```

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:countBy` | 顯示間隔（每 n 行顯示） |
| `w:start` | 起始值 |
| `w:restart` | 重新計數：`newPage`、`newSection`、`continuous` |
| `w:distance` | 與文字的距離（twips） |

---

## w:pgNumType（頁碼設定）

```xml
<w:pgNumType w:fmt="decimal"
             w:start="1"
             w:chapStyle="1"
             w:chapSep="hyphen"/>
```

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:fmt` | 編號格式 |
| `w:start` | 起始頁碼 |
| `w:chapStyle` | 章節標題樣式 |
| `w:chapSep` | 章節分隔符：`hyphen`、`period`、`colon`、`emDash`、`enDash` |

### w:fmt 格式

| 值 | 說明 | 範例 |
|----|------|------|
| `decimal` | 十進位 | 1, 2, 3 |
| `upperRoman` | 大寫羅馬 | I, II, III |
| `lowerRoman` | 小寫羅馬 | i, ii, iii |
| `upperLetter` | 大寫字母 | A, B, C |
| `lowerLetter` | 小寫字母 | a, b, c |
| `cardinalText` | 基數 | One, Two, Three |
| `ordinalText` | 序數 | First, Second |
| `numberInDash` | 短線數字 | -1-, -2- |
| `taiwaneseCounting` | 台灣計數 | 一, 二, 三 |

---

## w:docGrid（文件格線）

```xml
<w:docGrid w:type="lines"
           w:linePitch="360"
           w:charSpace="0"/>
```

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:type` | 格線類型 |
| `w:linePitch` | 行距（twips） |
| `w:charSpace` | 字元間距 |

### w:type 類型

| 值 | 說明 |
|----|------|
| `default` | 預設（無格線） |
| `lines` | 只有行格線 |
| `linesAndChars` | 行和字元格線 |
| `snapToChars` | 對齊字元格線 |

---

## w:headerReference / w:footerReference（頁首/頁尾參照）

```xml
<w:headerReference w:type="default" r:id="rId7"/>
<w:headerReference w:type="first" r:id="rId8"/>
<w:headerReference w:type="even" r:id="rId9"/>

<w:footerReference w:type="default" r:id="rId10"/>
<w:footerReference w:type="first" r:id="rId11"/>
<w:footerReference w:type="even" r:id="rId12"/>
```

### w:type 類型

| 值 | 說明 |
|----|------|
| `default` | 預設（奇數頁） |
| `first` | 首頁 |
| `even` | 偶數頁 |

### 需要配合的設定

```xml
<!-- 首頁不同 -->
<w:titlePg/>

<!-- 奇偶頁不同（在 settings.xml 中） -->
<w:evenAndOddHeaders/>
```

---

## w:titlePg（首頁不同）

啟用首頁不同的頁首/頁尾。

```xml
<w:sectPr>
    <w:headerReference w:type="first" r:id="rId1"/>
    <w:headerReference w:type="default" r:id="rId2"/>
    <w:footerReference w:type="first" r:id="rId3"/>
    <w:footerReference w:type="default" r:id="rId4"/>
    <w:titlePg/>
    <!-- ... -->
</w:sectPr>
```

---

## w:vAlign（垂直對齊）

頁面內容的垂直對齊。

```xml
<w:vAlign w:val="center"/>
```

| 值 | 說明 |
|----|------|
| `top` | 靠上（預設） |
| `center` | 置中 |
| `both` | 分散對齊 |
| `bottom` | 靠下 |

---

## w:textDirection（文字方向）

```xml
<w:textDirection w:val="tbRl"/>
```

| 值 | 說明 |
|----|------|
| `lrTb` | 左到右，上到下（預設） |
| `tbRl` | 上到下，右到左（直排） |
| `btLr` | 下到上，左到右 |
| `lrTbV` | 左到右，上到下（垂直） |
| `tbRlV` | 上到下，右到左（垂直） |
| `tbLrV` | 上到下，左到右（垂直） |

---

## 完整範例

### 標準 A4 文件

```xml
<w:sectPr>
    <!-- 頁面大小：A4 -->
    <w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>

    <!-- 頁邊距：標準 -->
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
             w:header="720" w:footer="720" w:gutter="0"/>

    <!-- 單欄 -->
    <w:cols w:space="720"/>

    <!-- 文件格線 -->
    <w:docGrid w:linePitch="360"/>
</w:sectPr>
```

### 橫向雙欄文件

```xml
<w:sectPr>
    <!-- 頁面大小：橫向 A4 -->
    <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>

    <!-- 頁邊距 -->
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
             w:header="720" w:footer="720" w:gutter="0"/>

    <!-- 雙欄 -->
    <w:cols w:num="2" w:space="720"/>
</w:sectPr>
```

### 帶頁首頁尾的文件

```xml
<w:sectPr>
    <!-- 頁首參照 -->
    <w:headerReference w:type="default" r:id="rId7"/>
    <w:headerReference w:type="first" r:id="rId8"/>

    <!-- 頁尾參照 -->
    <w:footerReference w:type="default" r:id="rId9"/>
    <w:footerReference w:type="first" r:id="rId10"/>

    <!-- 首頁不同 -->
    <w:titlePg/>

    <!-- 頁面設定 -->
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
             w:header="720" w:footer="720"/>
    <w:cols w:space="720"/>
    <w:docGrid w:linePitch="360"/>
</w:sectPr>
```

### 連續分節

```xml
<!-- 第一節結束 -->
<w:p>
    <w:pPr>
        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            <w:cols w:space="720"/>
            <w:type w:val="continuous"/>
        </w:sectPr>
    </w:pPr>
    <w:r><w:t>第一節最後一段</w:t></w:r>
</w:p>

<!-- 第二節開始（雙欄） -->
<w:p>
    <w:r><w:t>第二節內容（雙欄）</w:t></w:r>
</w:p>

<!-- 文件末節 -->
<w:sectPr>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    <w:cols w:num="2" w:space="720"/>
</w:sectPr>
```

---

## 下一步

- [42-headers-footers.md](42-headers-footers.md) - 頁首頁尾
- [43-page-numbers.md](43-page-numbers.md) - 頁碼設定
- [50-images.md](50-images.md) - 圖片
