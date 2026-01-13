# 段落 (Paragraph) 元素

## 概述

`w:p` 是 WordprocessingML 中最重要的元素之一，代表一個段落。

## 基本結構

```xml
<w:p>
    <w:pPr>
        <!-- 段落屬性 -->
    </w:pPr>
    <w:r>
        <!-- 文字運行 -->
    </w:r>
</w:p>
```

---

## w:p 子元素

| 元素 | 說明 | 必要 |
|------|------|------|
| `w:pPr` | 段落屬性 | 否 |
| `w:r` | 文字運行 (Run) | 否 |
| `w:hyperlink` | 超連結 | 否 |
| `w:bookmarkStart` | 書籤開始 | 否 |
| `w:bookmarkEnd` | 書籤結束 | 否 |
| `w:commentRangeStart` | 註解範圍開始 | 否 |
| `w:commentRangeEnd` | 註解範圍結束 | 否 |
| `w:fldSimple` | 簡單欄位 | 否 |
| `w:customXml` | 自訂 XML | 否 |
| `w:sdt` | 結構化文件標籤 | 否 |
| `w:smartTag` | 智慧標籤 | 否 |
| `w:proofErr` | 校對錯誤標記 | 否 |

---

## w:pPr（段落屬性）

### 完整屬性列表

| 元素 | 說明 | 範例值 |
|------|------|--------|
| `w:pStyle` | 段落樣式 | `Heading1`, `Normal` |
| `w:keepNext` | 與下段同頁 | - |
| `w:keepLines` | 段中不分頁 | - |
| `w:pageBreakBefore` | 段前分頁 | - |
| `w:framePr` | 文字框架屬性 | - |
| `w:widowControl` | 遺孤控制 | - |
| `w:numPr` | 編號屬性 | - |
| `w:suppressLineNumbers` | 隱藏行號 | - |
| `w:pBdr` | 段落邊框 | - |
| `w:shd` | 段落底色 | - |
| `w:tabs` | 定位點 | - |
| `w:suppressAutoHyphens` | 禁止自動連字號 | - |
| `w:kinsoku` | 禁則處理 | - |
| `w:wordWrap` | 自動換行 | - |
| `w:overflowPunct` | 標點溢出 | - |
| `w:topLinePunct` | 行首標點 | - |
| `w:autoSpaceDE` | 自動空格（拉丁與東亞） | - |
| `w:autoSpaceDN` | 自動空格（數字與東亞） | - |
| `w:bidi` | 雙向文字 | - |
| `w:adjustRightInd` | 調整右縮排 | - |
| `w:snapToGrid` | 對齊格線 | - |
| `w:spacing` | 間距設定 | - |
| `w:ind` | 縮排設定 | - |
| `w:contextualSpacing` | 相同樣式段落間距 | - |
| `w:mirrorIndents` | 鏡像縮排 | - |
| `w:suppressOverlap` | 禁止重疊 | - |
| `w:jc` | 對齊方式 | `left`, `center`, `right`, `both` |
| `w:textDirection` | 文字方向 | - |
| `w:textAlignment` | 文字對齊 | - |
| `w:textboxTightWrap` | 文字框緊密環繞 | - |
| `w:outlineLvl` | 大綱層級 | `0`-`9` |
| `w:divId` | HTML div ID | - |
| `w:cnfStyle` | 條件式格式 | - |
| `w:rPr` | 預設文字屬性 | - |
| `w:sectPr` | 分節屬性 | - |

---

## 常用段落屬性詳解

### w:pStyle（段落樣式）

```xml
<w:pPr>
    <w:pStyle w:val="Heading1"/>
</w:pPr>
```

常用樣式：
- `Normal` - 內文
- `Heading1` ~ `Heading9` - 標題 1-9
- `Title` - 標題
- `Subtitle` - 副標題
- `Quote` - 引言
- `ListParagraph` - 清單段落

### w:jc（對齊方式）

```xml
<w:pPr>
    <w:jc w:val="center"/>
</w:pPr>
```

| 值 | 說明 |
|----|------|
| `left` | 靠左對齊 |
| `center` | 置中對齊 |
| `right` | 靠右對齊 |
| `both` | 左右對齊（兩端對齊） |
| `distribute` | 分散對齊 |
| `mediumKashida` | 中等 Kashida（阿拉伯文） |
| `highKashida` | 高 Kashida |
| `lowKashida` | 低 Kashida |
| `thaiDistribute` | 泰文分散 |

### w:spacing（間距）

```xml
<w:pPr>
    <w:spacing w:before="240"      <!-- 段前間距（twips） -->
               w:after="120"       <!-- 段後間距（twips） -->
               w:line="360"        <!-- 行距值 -->
               w:lineRule="auto"/> <!-- 行距規則 -->
</w:pPr>
```

**w:lineRule 值：**

| 值 | 說明 |
|----|------|
| `auto` | 自動（w:line 以 1/240 行計） |
| `exact` | 固定（w:line 以 twips 計） |
| `atLeast` | 最小（w:line 以 twips 計） |

**常用行距換算：**

| 行距 | w:line 值 | w:lineRule |
|------|-----------|------------|
| 單倍行高 | 240 | auto |
| 1.5 倍行高 | 360 | auto |
| 雙倍行高 | 480 | auto |
| 固定 12pt | 240 | exact |
| 最小 12pt | 240 | atLeast |

### w:ind（縮排）

```xml
<w:pPr>
    <w:ind w:left="720"           <!-- 左縮排（twips） -->
           w:right="720"          <!-- 右縮排（twips） -->
           w:firstLine="720"      <!-- 首行縮排（twips） -->
           w:hanging="720"/>      <!-- 懸吊縮排（twips） -->
</w:pPr>
```

**注意：** `firstLine` 和 `hanging` 互斥，只能使用其中一個。

### w:numPr（編號/清單）

```xml
<w:pPr>
    <w:numPr>
        <w:ilvl w:val="0"/>    <!-- 清單層級（0-8） -->
        <w:numId w:val="1"/>   <!-- 編號定義 ID -->
    </w:numPr>
</w:pPr>
```

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

**邊框樣式 (w:val)：**
`none`, `single`, `thick`, `double`, `dotted`, `dashed`, `dotDash`, `dotDotDash`, `triple`, `thinThickSmallGap`, `thickThinSmallGap`, `thinThickThinSmallGap`, `thinThickMediumGap`, `thickThinMediumGap`, `thinThickThinMediumGap`, `thinThickLargeGap`, `thickThinLargeGap`, `thinThickThinLargeGap`, `wave`, `doubleWave`, `dashSmallGap`, `dashDotStroked`, `threeDEmboss`, `threeDEngrave`, `outset`, `inset`

### w:shd（段落底色）

```xml
<w:pPr>
    <w:shd w:val="clear"      <!-- 圖案樣式 -->
           w:color="auto"     <!-- 前景色 -->
           w:fill="FFFF00"/>  <!-- 背景色 -->
</w:pPr>
```

**圖案樣式 (w:val)：**
`clear`, `solid`, `horzStripe`, `vertStripe`, `reverseDiagStripe`, `diagStripe`, `horzCross`, `diagCross`, `thinHorzStripe`, `thinVertStripe`, `thinReverseDiagStripe`, `thinDiagStripe`, `thinHorzCross`, `thinDiagCross`, `pct5`, `pct10`, `pct12`, `pct15`, `pct20`, `pct25`, `pct30`, `pct35`, `pct37`, `pct40`, `pct45`, `pct50`, `pct55`, `pct60`, `pct62`, `pct65`, `pct70`, `pct75`, `pct80`, `pct85`, `pct87`, `pct90`, `pct95`

### w:tabs（定位點）

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

**定位點類型 (w:val)：**
- `left` - 靠左
- `center` - 置中
- `right` - 靠右
- `decimal` - 小數點對齊
- `bar` - 垂直線
- `clear` - 清除
- `num` - 編號

**前導字元 (w:leader)：**
- `none` - 無
- `dot` - 點
- `hyphen` - 連字號
- `underscore` - 底線
- `heavy` - 粗底線
- `middleDot` - 中間點

---

## 完整範例

### 標題段落

```xml
<w:p>
    <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:keepNext/>
        <w:keepLines/>
        <w:spacing w:before="480" w:after="120"/>
        <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:bookmarkStart w:id="0" w:name="_Toc123456789"/>
    <w:r>
        <w:t>第一章 簡介</w:t>
    </w:r>
    <w:bookmarkEnd w:id="0"/>
</w:p>
```

### 首行縮排段落

```xml
<w:p>
    <w:pPr>
        <w:ind w:firstLine="480"/>
        <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
    </w:pPr>
    <w:r>
        <w:t>這是一個首行縮排的段落。首行會自動縮進兩個字元的寬度，讓文章看起來更加整齊美觀。</w:t>
    </w:r>
</w:p>
```

### 引言段落

```xml
<w:p>
    <w:pPr>
        <w:pStyle w:val="Quote"/>
        <w:ind w:left="720" w:right="720"/>
        <w:jc w:val="both"/>
        <w:pBdr>
            <w:left w:val="single" w:sz="18" w:space="12" w:color="999999"/>
        </w:pBdr>
        <w:shd w:val="clear" w:fill="F5F5F5"/>
    </w:pPr>
    <w:r>
        <w:rPr>
            <w:i/>
        </w:rPr>
        <w:t>「知識就是力量。」——培根</w:t>
    </w:r>
</w:p>
```

### 編號清單項目

```xml
<w:p>
    <w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:numPr>
            <w:ilvl w:val="0"/>
            <w:numId w:val="1"/>
        </w:numPr>
        <w:spacing w:after="0"/>
    </w:pPr>
    <w:r>
        <w:t>第一個清單項目</w:t>
    </w:r>
</w:p>
<w:p>
    <w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:numPr>
            <w:ilvl w:val="1"/>
            <w:numId w:val="1"/>
        </w:numPr>
        <w:spacing w:after="0"/>
    </w:pPr>
    <w:r>
        <w:t>巢狀項目</w:t>
    </w:r>
</w:p>
```

---

## 下一步

- [12-run.md](12-run.md) - 文字運行 (Run) 元素
- [13-text-formatting.md](13-text-formatting.md) - 文字格式化屬性
- [14-paragraph-formatting.md](14-paragraph-formatting.md) - 段落格式化詳解
