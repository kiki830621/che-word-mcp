# 段落格式化屬性詳解

## 概述

`w:pPr` (Paragraph Properties) 包含所有段落級別的格式設定。本文件詳細說明每個屬性的用法。

---

## 段落樣式

### w:pStyle（段落樣式）

```xml
<w:pPr>
    <w:pStyle w:val="Heading1"/>
</w:pPr>
```

#### 內建樣式 ID

| 類別 | 樣式 ID | 說明 |
|------|---------|------|
| **標題** | `Title` | 標題 |
| | `Subtitle` | 副標題 |
| | `Heading1` ~ `Heading9` | 標題 1-9 |
| **內文** | `Normal` | 內文 |
| | `BodyText` | 本文 |
| | `BodyTextIndent` | 本文縮排 |
| | `BodyText2` | 本文 2 |
| **清單** | `ListParagraph` | 清單段落 |
| | `ListBullet` | 項目符號清單 |
| | `ListNumber` | 編號清單 |
| **引言** | `Quote` | 引言 |
| | `IntenseQuote` | 強調引言 |
| **其他** | `NoSpacing` | 無間距 |
| | `TOCHeading` | 目錄標題 |
| | `TOC1` ~ `TOC9` | 目錄層級 |
| | `Caption` | 圖表標號 |
| | `Footnote Text` | 腳註文字 |
| | `Endnote Text` | 尾註文字 |
| | `Header` | 頁首 |
| | `Footer` | 頁尾 |

---

## 對齊方式

### w:jc（水平對齊）

```xml
<w:pPr>
    <w:jc w:val="center"/>
</w:pPr>
```

| 值 | 說明 | 適用 |
|----|------|------|
| `left` | 靠左對齊 | 預設（LTR 文字） |
| `center` | 置中對齊 | 標題常用 |
| `right` | 靠右對齊 | 預設（RTL 文字） |
| `both` | 左右對齊 | 內文常用 |
| `distribute` | 分散對齊 | 東亞文字 |
| `start` | 文字方向起始端 | |
| `end` | 文字方向結束端 | |
| `mediumKashida` | 中等 Kashida | 阿拉伯文 |
| `highKashida` | 高 Kashida | 阿拉伯文 |
| `lowKashida` | 低 Kashida | 阿拉伯文 |
| `thaiDistribute` | 泰文分散 | 泰文 |

---

## 縮排

### w:ind（縮排設定）

所有值以 **twips** 為單位（1 inch = 1440 twips, 1 cm ≈ 567 twips）。

```xml
<w:pPr>
    <w:ind w:left="720"           <!-- 左縮排 -->
           w:right="720"          <!-- 右縮排 -->
           w:firstLine="480"      <!-- 首行縮排 -->
           w:hanging="480"/>      <!-- 懸吊縮排 -->
</w:pPr>
```

| 屬性 | 說明 | 範例 |
|------|------|------|
| `w:left` | 左縮排 | 720 = 0.5 inch |
| `w:right` | 右縮排 | 720 = 0.5 inch |
| `w:firstLine` | 首行縮排 | 480 = 約2個中文字 |
| `w:hanging` | 懸吊縮排 | 首行以外縮排 |
| `w:leftChars` | 左縮排（字元數） | 100 = 1字元 |
| `w:rightChars` | 右縮排（字元數） | |
| `w:firstLineChars` | 首行縮排（字元數） | 200 = 2字元 |
| `w:hangingChars` | 懸吊縮排（字元數） | |

**注意：** `firstLine` 和 `hanging` 互斥，只能使用其中一個。

#### 常用縮排換算

| 縮排量 | twips 值 |
|--------|----------|
| 0.25 inch | 360 |
| 0.5 inch | 720 |
| 1 inch | 1440 |
| 1 cm | 567 |
| 2 字元 | ~480 |

#### 首行縮排範例

```xml
<!-- 首行縮排 2 字元 -->
<w:pPr>
    <w:ind w:firstLine="480"/>
</w:pPr>

<!-- 首行縮排 2 字元（使用字元單位） -->
<w:pPr>
    <w:ind w:firstLineChars="200"/>
</w:pPr>
```

#### 懸吊縮排範例（清單樣式）

```xml
<w:pPr>
    <w:ind w:left="720" w:hanging="360"/>
</w:pPr>
```

---

## 間距

### w:spacing（段落間距）

```xml
<w:pPr>
    <w:spacing w:before="240"      <!-- 段前間距 (twips) -->
               w:beforeLines="50"  <!-- 段前間距 (行數×100) -->
               w:beforeAutospacing="1"  <!-- 自動段前 -->
               w:after="120"       <!-- 段後間距 (twips) -->
               w:afterLines="50"   <!-- 段後間距 (行數×100) -->
               w:afterAutospacing="1"   <!-- 自動段後 -->
               w:line="360"        <!-- 行距值 -->
               w:lineRule="auto"/> <!-- 行距規則 -->
</w:pPr>
```

### 段前/段後間距

| 屬性 | 單位 | 說明 |
|------|------|------|
| `w:before` | twips | 段前間距 |
| `w:after` | twips | 段後間距 |
| `w:beforeLines` | 行數×100 | 段前（50 = 0.5行） |
| `w:afterLines` | 行數×100 | 段後（50 = 0.5行） |
| `w:beforeAutospacing` | 布林 | 自動段前（HTML 模式） |
| `w:afterAutospacing` | 布林 | 自動段後（HTML 模式） |

#### 常用段落間距

| 間距 | twips 值 |
|------|----------|
| 6pt | 120 |
| 12pt | 240 |
| 18pt | 360 |
| 24pt | 480 |

### 行距設定

```xml
<!-- 單倍行高 -->
<w:spacing w:line="240" w:lineRule="auto"/>

<!-- 1.5 倍行高 -->
<w:spacing w:line="360" w:lineRule="auto"/>

<!-- 雙倍行高 -->
<w:spacing w:line="480" w:lineRule="auto"/>

<!-- 固定 24pt 行距 -->
<w:spacing w:line="480" w:lineRule="exact"/>

<!-- 最小 18pt 行距 -->
<w:spacing w:line="360" w:lineRule="atLeast"/>
```

### w:lineRule 行距規則

| 值 | 說明 | w:line 單位 |
|----|------|-------------|
| `auto` | 自動（倍數） | 1/240 行（240=單倍） |
| `exact` | 固定行距 | twips |
| `atLeast` | 最小行距 | twips |

#### 行距換算表

| 行距類型 | w:lineRule | w:line 值 |
|----------|------------|-----------|
| 單倍行高 | auto | 240 |
| 1.15 倍 | auto | 276 |
| 1.5 倍 | auto | 360 |
| 雙倍行高 | auto | 480 |
| 2.5 倍 | auto | 600 |
| 3 倍 | auto | 720 |
| 固定 12pt | exact | 240 |
| 固定 14pt | exact | 280 |
| 固定 16pt | exact | 320 |
| 固定 18pt | exact | 360 |
| 最小 12pt | atLeast | 240 |

### w:contextualSpacing（相同樣式間距）

當前後段落使用相同樣式時，自動移除段落間距。

```xml
<w:pPr>
    <w:contextualSpacing/>
</w:pPr>
```

---

## 段落邊框

### w:pBdr（段落邊框）

```xml
<w:pPr>
    <w:pBdr>
        <w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="4" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="4" w:color="000000"/>
        <w:between w:val="single" w:sz="4" w:space="1" w:color="000000"/>
        <w:bar w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:pBdr>
</w:pPr>
```

### 邊框元素

| 元素 | 說明 |
|------|------|
| `w:top` | 上邊框 |
| `w:left` | 左邊框 |
| `w:bottom` | 下邊框 |
| `w:right` | 右邊框 |
| `w:between` | 相同邊框設定段落之間的分隔線 |
| `w:bar` | 頁面邊緣的垂直線 |

### 邊框屬性

| 屬性 | 說明 |
|------|------|
| `w:val` | 邊框樣式 |
| `w:sz` | 邊框寬度（1/8 pt） |
| `w:space` | 與文字的間距（pt） |
| `w:color` | 顏色（RGB 或 auto） |
| `w:themeColor` | 主題色彩 |
| `w:shadow` | 陰影效果 |
| `w:frame` | 框架效果 |

### 邊框樣式 (w:val)

| 值 | 說明 |
|----|------|
| `none` | 無邊框 |
| `nil` | 移除繼承的邊框 |
| `single` | 單線 |
| `thick` | 粗線 |
| `double` | 雙線 |
| `dotted` | 點線 |
| `dashed` | 虛線 |
| `dotDash` | 點虛線 |
| `dotDotDash` | 雙點虛線 |
| `triple` | 三線 |
| `wave` | 波浪線 |
| `doubleWave` | 雙波浪線 |
| `thinThickSmallGap` | 細粗細（小間距） |
| `thickThinSmallGap` | 粗細粗（小間距） |
| `thinThickThinSmallGap` | 細粗細（小間距） |
| `thinThickMediumGap` | 細粗（中間距） |
| `thickThinMediumGap` | 粗細（中間距） |
| `thinThickThinMediumGap` | 細粗細（中間距） |
| `thinThickLargeGap` | 細粗（大間距） |
| `thickThinLargeGap` | 粗細（大間距） |
| `thinThickThinLargeGap` | 細粗細（大間距） |
| `threeDEmboss` | 3D 浮凸 |
| `threeDEngrave` | 3D 雕刻 |
| `outset` | 外凸 |
| `inset` | 內凹 |
| `dashSmallGap` | 小間距虛線 |
| `dashDotStroked` | 筆畫點虛線 |

### 引言邊框範例

```xml
<w:pPr>
    <w:pBdr>
        <w:left w:val="single" w:sz="24" w:space="12" w:color="4472C4"/>
    </w:pBdr>
    <w:ind w:left="720"/>
    <w:shd w:val="clear" w:fill="F2F2F2"/>
</w:pPr>
```

---

## 段落底色

### w:shd（段落底色）

```xml
<w:pPr>
    <w:shd w:val="clear"          <!-- 圖案樣式 -->
           w:color="auto"         <!-- 前景色 -->
           w:fill="FFFF00"/>      <!-- 背景色 -->
</w:pPr>
```

### 圖案樣式 (w:val)

| 值 | 說明 | 效果 |
|----|------|------|
| `clear` | 無圖案 | 純色背景 |
| `solid` | 實心 | 使用前景色 |
| `pct5` ~ `pct95` | 百分比 | 5%-95% 填充 |
| `horzStripe` | 水平條紋 | |
| `vertStripe` | 垂直條紋 | |
| `diagStripe` | 對角條紋 | |
| `reverseDiagStripe` | 反對角條紋 | |
| `horzCross` | 水平交叉 | |
| `diagCross` | 對角交叉 | |
| `thinHorzStripe` | 細水平條紋 | |
| `thinVertStripe` | 細垂直條紋 | |
| `thinDiagStripe` | 細對角條紋 | |
| `thinReverseDiagStripe` | 細反對角條紋 | |
| `thinHorzCross` | 細水平交叉 | |
| `thinDiagCross` | 細對角交叉 | |

---

## 定位點

### w:tabs（定位點設定）

```xml
<w:pPr>
    <w:tabs>
        <w:tab w:val="left" w:pos="2880"/>
        <w:tab w:val="center" w:pos="4680"/>
        <w:tab w:val="right" w:pos="9360"/>
        <w:tab w:val="decimal" w:pos="7200"/>
        <w:tab w:val="bar" w:pos="5760"/>
        <w:tab w:val="clear" w:pos="720"/>
    </w:tabs>
</w:pPr>
```

### 定位點類型 (w:val)

| 值 | 說明 | 效果 |
|----|------|------|
| `left` | 靠左 | 文字從定位點開始向右 |
| `center` | 置中 | 文字在定位點位置置中 |
| `right` | 靠右 | 文字從定位點結束向左 |
| `decimal` | 小數點對齊 | 數字小數點對齊 |
| `bar` | 垂直線 | 在定位點位置顯示垂直線 |
| `clear` | 清除 | 清除指定位置的定位點 |
| `num` | 編號 | 用於編號清單 |

### 前導字元 (w:leader)

```xml
<w:tab w:val="right" w:pos="9360" w:leader="dot"/>
```

| 值 | 說明 | 效果 |
|----|------|------|
| `none` | 無前導字元 | |
| `dot` | 點 | .......... |
| `hyphen` | 連字號 | ---------- |
| `underscore` | 底線 | __________ |
| `heavy` | 粗底線 | ━━━━━━━━━━ |
| `middleDot` | 中間點 | ·········· |

### 目錄前導字元範例

```xml
<w:pPr>
    <w:tabs>
        <w:tab w:val="right" w:leader="dot" w:pos="9360"/>
    </w:tabs>
</w:pPr>
```

---

## 編號與清單

### w:numPr（編號屬性）

```xml
<w:pPr>
    <w:numPr>
        <w:ilvl w:val="0"/>    <!-- 清單層級（0-8） -->
        <w:numId w:val="1"/>   <!-- 編號定義 ID -->
    </w:numPr>
</w:pPr>
```

| 元素 | 說明 |
|------|------|
| `w:ilvl` | 清單層級（0 = 第一層，最多 9 層） |
| `w:numId` | 參照 numbering.xml 中的編號定義 |

#### 多層清單範例

```xml
<!-- 第一層 -->
<w:p>
    <w:pPr>
        <w:numPr>
            <w:ilvl w:val="0"/>
            <w:numId w:val="1"/>
        </w:numPr>
    </w:pPr>
    <w:r><w:t>第一項</w:t></w:r>
</w:p>

<!-- 第二層 -->
<w:p>
    <w:pPr>
        <w:numPr>
            <w:ilvl w:val="1"/>
            <w:numId w:val="1"/>
        </w:numPr>
    </w:pPr>
    <w:r><w:t>子項目</w:t></w:r>
</w:p>
```

詳見：[32-numbering.md](32-numbering.md)

---

## 分頁控制

### w:keepNext（與下段同頁）

確保此段落與下一段落在同一頁。

```xml
<w:pPr>
    <w:keepNext/>
</w:pPr>
```

**用途：** 標題與內文不分離

### w:keepLines（段中不分頁）

確保段落中的所有行在同一頁。

```xml
<w:pPr>
    <w:keepLines/>
</w:pPr>
```

**用途：** 避免段落被分頁符切割

### w:pageBreakBefore（段前分頁）

在此段落前插入分頁符。

```xml
<w:pPr>
    <w:pageBreakBefore/>
</w:pPr>
```

**用途：** 章節標題強制從新頁開始

### w:widowControl（遺孤控制）

避免段落首行或末行單獨出現在頁面。

```xml
<w:pPr>
    <w:widowControl/>
</w:pPr>

<!-- 停用遺孤控制 -->
<w:pPr>
    <w:widowControl w:val="0"/>
</w:pPr>
```

---

## 大綱層級

### w:outlineLvl（大綱層級）

定義段落在文件大綱中的層級，用於目錄生成。

```xml
<w:pPr>
    <w:outlineLvl w:val="0"/>  <!-- 層級 1（最高） -->
</w:pPr>
```

| 值 | 說明 |
|----|------|
| `0` | 層級 1（如 Heading1） |
| `1` | 層級 2（如 Heading2） |
| ... | ... |
| `8` | 層級 9（如 Heading9） |
| `9` | Body Text（不納入大綱） |

---

## 文字方向

### w:textDirection（文字方向）

```xml
<w:pPr>
    <w:textDirection w:val="tbRl"/>
</w:pPr>
```

| 值 | 說明 |
|----|------|
| `lrTb` | 左到右，上到下（預設） |
| `tbRl` | 上到下，右到左（直排） |
| `btLr` | 下到上，左到右 |
| `lrTbV` | 垂直：左到右，上到下 |
| `tbRlV` | 垂直：上到下，右到左 |
| `tbLrV` | 垂直：上到下，左到右 |

### w:bidi（雙向文字）

標記段落為從右到左方向。

```xml
<w:pPr>
    <w:bidi/>
</w:pPr>
```

---

## 東亞排版

### w:kinsoku（禁則處理）

啟用東亞文字的禁則處理（避免標點在行首/行尾）。

```xml
<w:pPr>
    <w:kinsoku/>
</w:pPr>
```

### w:wordWrap（自動換行）

控制東亞文字的換行方式。

```xml
<w:pPr>
    <w:wordWrap w:val="0"/>  <!-- 禁止在單字中間換行 -->
</w:pPr>
```

### w:overflowPunct（標點溢出）

允許標點符號溢出頁邊距。

```xml
<w:pPr>
    <w:overflowPunct/>
</w:pPr>
```

### w:topLinePunct（行首標點壓縮）

壓縮行首的標點符號。

```xml
<w:pPr>
    <w:topLinePunct/>
</w:pPr>
```

### w:autoSpaceDE（自動空格：拉丁與東亞）

在拉丁文字和東亞文字之間自動插入空格。

```xml
<w:pPr>
    <w:autoSpaceDE/>
</w:pPr>
```

### w:autoSpaceDN（自動空格：數字與東亞）

在數字和東亞文字之間自動插入空格。

```xml
<w:pPr>
    <w:autoSpaceDN/>
</w:pPr>
```

---

## 其他屬性

### w:snapToGrid（對齊格線）

將段落行對齊文件格線。

```xml
<w:pPr>
    <w:snapToGrid/>
</w:pPr>
```

### w:suppressLineNumbers（隱藏行號）

不顯示此段落的行號。

```xml
<w:pPr>
    <w:suppressLineNumbers/>
</w:pPr>
```

### w:suppressAutoHyphens（禁止自動連字號）

禁止自動在此段落中插入連字號。

```xml
<w:pPr>
    <w:suppressAutoHyphens/>
</w:pPr>
```

### w:mirrorIndents（鏡像縮排）

在雙面印刷中，奇偶頁互換左右縮排。

```xml
<w:pPr>
    <w:mirrorIndents/>
</w:pPr>
```

### w:adjustRightInd（調整右縮排）

自動調整右縮排以對齊格線。

```xml
<w:pPr>
    <w:adjustRightInd/>
</w:pPr>
```

---

## 框架屬性

### w:framePr（文字框架）

將段落放入浮動框架中。

```xml
<w:pPr>
    <w:framePr w:w="2880"           <!-- 寬度 (twips) -->
               w:h="1440"           <!-- 高度 (twips) -->
               w:hRule="exact"      <!-- 高度規則 -->
               w:wrap="around"      <!-- 文繞圖方式 -->
               w:vAnchor="text"     <!-- 垂直錨點 -->
               w:hAnchor="page"     <!-- 水平錨點 -->
               w:x="1440"           <!-- 水平位置 -->
               w:y="720"            <!-- 垂直位置 -->
               w:xAlign="center"    <!-- 水平對齊 -->
               w:yAlign="top"/>     <!-- 垂直對齊 -->
</w:pPr>
```

| 屬性 | 說明 |
|------|------|
| `w:w` | 框架寬度 (twips) |
| `w:h` | 框架高度 (twips) |
| `w:hRule` | 高度規則：`auto`, `exact`, `atLeast` |
| `w:wrap` | 文繞圖：`around`, `notBeside`, `none`, `through`, `tight` |
| `w:vAnchor` | 垂直錨點：`text`, `margin`, `page` |
| `w:hAnchor` | 水平錨點：`text`, `margin`, `page` |
| `w:x` | 水平位置偏移 |
| `w:y` | 垂直位置偏移 |
| `w:xAlign` | 水平對齊：`left`, `center`, `right`, `inside`, `outside` |
| `w:yAlign` | 垂直對齊：`top`, `center`, `bottom`, `inside`, `outside` |

---

## 段落預設文字屬性

### w:rPr（段落預設文字屬性）

定義段落中文字的預設格式。

```xml
<w:pPr>
    <w:rPr>
        <w:rFonts w:ascii="Arial" w:eastAsia="微軟正黑體"/>
        <w:sz w:val="24"/>
        <w:color w:val="000000"/>
    </w:rPr>
</w:pPr>
```

**注意：** 這會影響段落標記和段落中沒有明確格式的文字。

---

## 分節屬性

### w:sectPr（分節屬性）

當 `w:sectPr` 出現在 `w:pPr` 中時，表示此段落是一個節的最後一段。

```xml
<w:p>
    <w:pPr>
        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            <w:type w:val="continuous"/>
        </w:sectPr>
    </w:pPr>
    <w:r><w:t>節的最後一段</w:t></w:r>
</w:p>
```

詳見：[40-section.md](40-section.md)

---

## 完整範例

### 標題段落

```xml
<w:p>
    <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:keepNext/>
        <w:keepLines/>
        <w:spacing w:before="480" w:after="240"/>
        <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:r>
        <w:t>第一章 簡介</w:t>
    </w:r>
</w:p>
```

### 引言段落

```xml
<w:p>
    <w:pPr>
        <w:pStyle w:val="Quote"/>
        <w:ind w:left="720" w:right="720"/>
        <w:pBdr>
            <w:left w:val="single" w:sz="24" w:space="12" w:color="4472C4"/>
        </w:pBdr>
        <w:shd w:val="clear" w:fill="F2F2F2"/>
        <w:spacing w:before="240" w:after="240"/>
        <w:jc w:val="both"/>
    </w:pPr>
    <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>「學而不思則罔，思而不學則殆。」——《論語》</w:t>
    </w:r>
</w:p>
```

### 首行縮排段落

```xml
<w:p>
    <w:pPr>
        <w:ind w:firstLine="480"/>
        <w:spacing w:line="360" w:lineRule="auto" w:after="200"/>
        <w:jc w:val="both"/>
    </w:pPr>
    <w:r>
        <w:t>這是一個首行縮排的段落，常用於中文內文排版。首行會自動縮進兩個字元的寬度，讓文章看起來更加整齊美觀。</w:t>
    </w:r>
</w:p>
```

---

## 屬性順序

`w:pPr` 中的子元素應按以下順序排列（建議但非強制）：

1. `w:pStyle`
2. `w:keepNext`
3. `w:keepLines`
4. `w:pageBreakBefore`
5. `w:framePr`
6. `w:widowControl`
7. `w:numPr`
8. `w:suppressLineNumbers`
9. `w:pBdr`
10. `w:shd`
11. `w:tabs`
12. `w:suppressAutoHyphens`
13. `w:kinsoku`
14. `w:wordWrap`
15. `w:overflowPunct`
16. `w:topLinePunct`
17. `w:autoSpaceDE`
18. `w:autoSpaceDN`
19. `w:bidi`
20. `w:adjustRightInd`
21. `w:snapToGrid`
22. `w:spacing`
23. `w:ind`
24. `w:contextualSpacing`
25. `w:mirrorIndents`
26. `w:suppressOverlap`
27. `w:jc`
28. `w:textDirection`
29. `w:textAlignment`
30. `w:textboxTightWrap`
31. `w:outlineLvl`
32. `w:divId`
33. `w:cnfStyle`
34. `w:rPr`
35. `w:sectPr`

---

## 下一步

- [20-table.md](20-table.md) - 表格結構
- [30-styles.md](30-styles.md) - 樣式系統
- [40-section.md](40-section.md) - 分節屬性
