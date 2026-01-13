# 文字運行 (Run) 元素

## 概述

`w:r` (Run) 是段落中的文字容器，代表一段具有相同格式的文字。

## 基本結構

```xml
<w:r>
    <w:rPr>
        <!-- 文字屬性 -->
    </w:rPr>
    <w:t>文字內容</w:t>
</w:r>
```

---

## w:r 子元素

### 內容元素

| 元素 | 說明 |
|------|------|
| `w:t` | 文字 |
| `w:tab` | 定位點字元 |
| `w:br` | 換行/分頁符 |
| `w:cr` | 歸位字元 |
| `w:sym` | 符號 |
| `w:softHyphen` | 軟連字號 |
| `w:noBreakHyphen` | 不斷行連字號 |
| `w:drawing` | 繪圖（圖片） |
| `w:object` | 嵌入物件 |
| `w:pict` | VML 圖形 |
| `w:fldChar` | 欄位字元 |
| `w:instrText` | 欄位指令 |
| `w:delText` | 刪除的文字（追蹤修訂） |

### 屬性元素

| 元素 | 說明 |
|------|------|
| `w:rPr` | 文字運行屬性 |

### 參照元素

| 元素 | 說明 |
|------|------|
| `w:footnoteReference` | 腳註參照 |
| `w:endnoteReference` | 尾註參照 |
| `w:commentReference` | 註解參照 |
| `w:annotationRef` | 註釋參照 |

---

## w:t（文字）

### 屬性

| 屬性 | 說明 |
|------|------|
| `xml:space` | 空白處理方式 |

```xml
<!-- 保留空白 -->
<w:t xml:space="preserve">Hello    World</w:t>

<!-- 正常處理（預設） -->
<w:t>Hello World</w:t>
```

### 重要：空白處理

如果文字開頭或結尾有空白，或包含連續空白，必須使用 `xml:space="preserve"`，否則空白會被忽略。

---

## w:br（換行/分頁）

```xml
<!-- 軟換行（Shift+Enter） -->
<w:br/>

<!-- 分頁符 -->
<w:br w:type="page"/>

<!-- 分欄符 -->
<w:br w:type="column"/>

<!-- 文字環繞換行 -->
<w:br w:type="textWrapping" w:clear="all"/>
```

### w:type 值

| 值 | 說明 |
|----|------|
| `page` | 分頁符 |
| `column` | 分欄符 |
| `textWrapping` | 文字環繞換行 |

### w:clear 值（用於 textWrapping）

| 值 | 說明 |
|----|------|
| `none` | 無 |
| `left` | 清除左側 |
| `right` | 清除右側 |
| `all` | 清除兩側 |

---

## w:rPr（文字運行屬性）

### 完整屬性列表

| 元素 | 說明 | 範例 |
|------|------|------|
| `w:rStyle` | 字元樣式 | `Emphasis` |
| `w:rFonts` | 字型 | - |
| `w:b` | 粗體 | - |
| `w:bCs` | 複雜字型粗體 | - |
| `w:i` | 斜體 | - |
| `w:iCs` | 複雜字型斜體 | - |
| `w:caps` | 全部大寫 | - |
| `w:smallCaps` | 小型大寫 | - |
| `w:strike` | 刪除線 | - |
| `w:dstrike` | 雙刪除線 | - |
| `w:outline` | 空心字 | - |
| `w:shadow` | 陰影 | - |
| `w:emboss` | 浮凸 | - |
| `w:imprint` | 陰刻 | - |
| `w:noProof` | 不校對 | - |
| `w:snapToGrid` | 對齊格線 | - |
| `w:vanish` | 隱藏文字 | - |
| `w:webHidden` | 網頁隱藏 | - |
| `w:color` | 文字顏色 | - |
| `w:spacing` | 字元間距 | - |
| `w:w` | 字元縮放 | - |
| `w:kern` | 字距調整 | - |
| `w:position` | 垂直位置 | - |
| `w:sz` | 字型大小 | - |
| `w:szCs` | 複雜字型大小 | - |
| `w:highlight` | 螢光筆顏色 | - |
| `w:u` | 底線 | - |
| `w:effect` | 文字效果 | - |
| `w:bdr` | 文字邊框 | - |
| `w:shd` | 文字底色 | - |
| `w:fitText` | 調整文字寬度 | - |
| `w:vertAlign` | 上下標 | - |
| `w:rtl` | 從右到左 | - |
| `w:cs` | 複雜字型 | - |
| `w:em` | 強調符號 | - |
| `w:lang` | 語言 | - |
| `w:eastAsianLayout` | 東亞版面配置 | - |
| `w:specVanish` | 特殊隱藏 | - |

---

## 常用屬性詳解

### w:rFonts（字型）

```xml
<w:rPr>
    <w:rFonts w:ascii="Arial"           <!-- 拉丁文字 -->
              w:hAnsi="Arial"           <!-- 高位 ANSI -->
              w:eastAsia="新細明體"      <!-- 東亞文字 -->
              w:cs="Arial"/>            <!-- 複雜字型 -->
</w:rPr>
```

### w:b, w:i（粗體、斜體）

```xml
<w:rPr>
    <w:b/>      <!-- 粗體 -->
    <w:i/>      <!-- 斜體 -->
</w:rPr>

<!-- 取消粗體（在有繼承的情況下） -->
<w:rPr>
    <w:b w:val="0"/>
</w:rPr>
```

### w:sz（字型大小）

字型大小以**半點**為單位。

```xml
<w:rPr>
    <w:sz w:val="24"/>   <!-- 12pt (24 半點) -->
    <w:szCs w:val="24"/> <!-- 複雜字型 12pt -->
</w:rPr>
```

**換算：** `w:val = 字型大小(pt) × 2`

| 字型大小 | w:val |
|----------|-------|
| 9pt | 18 |
| 10pt | 20 |
| 10.5pt | 21 |
| 11pt | 22 |
| 12pt | 24 |
| 14pt | 28 |
| 16pt | 32 |
| 18pt | 36 |
| 20pt | 40 |
| 24pt | 48 |
| 36pt | 72 |
| 48pt | 96 |

### w:color（文字顏色）

```xml
<w:rPr>
    <w:color w:val="FF0000"/>        <!-- 紅色 (RGB Hex) -->
    <w:color w:val="auto"/>          <!-- 自動（通常是黑色） -->
    <w:color w:themeColor="accent1"/> <!-- 佈景主題色彩 -->
</w:rPr>
```

### w:highlight（螢光筆）

```xml
<w:rPr>
    <w:highlight w:val="yellow"/>
</w:rPr>
```

**可用顏色：**
`black`, `blue`, `cyan`, `darkBlue`, `darkCyan`, `darkGray`, `darkGreen`, `darkMagenta`, `darkRed`, `darkYellow`, `green`, `lightGray`, `magenta`, `red`, `white`, `yellow`

### w:u（底線）

```xml
<w:rPr>
    <w:u w:val="single"/>              <!-- 單底線 -->
    <w:u w:val="double"/>              <!-- 雙底線 -->
    <w:u w:val="single" w:color="FF0000"/> <!-- 紅色底線 -->
</w:rPr>
```

**底線樣式 (w:val)：**
`single`, `words`, `double`, `thick`, `dotted`, `dottedHeavy`, `dash`, `dashedHeavy`, `dashLong`, `dashLongHeavy`, `dotDash`, `dashDotHeavy`, `dotDotDash`, `dashDotDotHeavy`, `wave`, `wavyHeavy`, `wavyDouble`, `none`

### w:strike, w:dstrike（刪除線）

```xml
<w:rPr>
    <w:strike/>   <!-- 單刪除線 -->
</w:rPr>

<w:rPr>
    <w:dstrike/>  <!-- 雙刪除線 -->
</w:rPr>
```

### w:vertAlign（上下標）

```xml
<w:rPr>
    <w:vertAlign w:val="superscript"/>  <!-- 上標 -->
</w:rPr>

<w:rPr>
    <w:vertAlign w:val="subscript"/>    <!-- 下標 -->
</w:rPr>
```

### w:spacing（字元間距）

```xml
<w:rPr>
    <w:spacing w:val="20"/>   <!-- 擴展 1pt (以 twips 計) -->
    <w:spacing w:val="-20"/>  <!-- 壓縮 1pt -->
</w:rPr>
```

### w:w（字元縮放）

```xml
<w:rPr>
    <w:w w:val="150"/>  <!-- 150% 寬度 -->
    <w:w w:val="50"/>   <!-- 50% 寬度 -->
</w:rPr>
```

### w:position（垂直位置）

```xml
<w:rPr>
    <w:position w:val="6"/>   <!-- 上移 3pt (以半點計) -->
    <w:position w:val="-6"/>  <!-- 下移 3pt -->
</w:rPr>
```

### w:shd（文字底色）

```xml
<w:rPr>
    <w:shd w:val="clear" w:fill="FFFF00"/>  <!-- 黃色背景 -->
</w:rPr>
```

### w:bdr（文字邊框）

```xml
<w:rPr>
    <w:bdr w:val="single" w:sz="4" w:space="1" w:color="000000"/>
</w:rPr>
```

---

## 欄位代碼

欄位使用三個元素表示：開始、指令、結束。

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
    <w:t>1</w:t>  <!-- 欄位值 -->
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### 常用欄位代碼

| 欄位 | 說明 |
|------|------|
| `PAGE` | 目前頁碼 |
| `NUMPAGES` | 總頁數 |
| `DATE` | 日期 |
| `TIME` | 時間 |
| `AUTHOR` | 作者 |
| `TITLE` | 標題 |
| `TOC` | 目錄 |
| `REF` | 交互參照 |
| `HYPERLINK` | 超連結 |
| `SEQ` | 序號 |

---

## 完整範例

### 混合格式文字

```xml
<w:p>
    <w:r>
        <w:t>這是</w:t>
    </w:r>
    <w:r>
        <w:rPr>
            <w:b/>
        </w:rPr>
        <w:t>粗體</w:t>
    </w:r>
    <w:r>
        <w:t>和</w:t>
    </w:r>
    <w:r>
        <w:rPr>
            <w:i/>
            <w:color w:val="FF0000"/>
        </w:rPr>
        <w:t>紅色斜體</w:t>
    </w:r>
    <w:r>
        <w:t>文字。</w:t>
    </w:r>
</w:p>
```

### 數學公式

```xml
<w:p>
    <w:r>
        <w:t>E = mc</w:t>
    </w:r>
    <w:r>
        <w:rPr>
            <w:vertAlign w:val="superscript"/>
        </w:rPr>
        <w:t>2</w:t>
    </w:r>
</w:p>
```

### 腳註參照

```xml
<w:p>
    <w:r>
        <w:t>根據研究</w:t>
    </w:r>
    <w:r>
        <w:rPr>
            <w:rStyle w:val="FootnoteReference"/>
        </w:rPr>
        <w:footnoteReference w:id="1"/>
    </w:r>
    <w:r>
        <w:t>顯示...</w:t>
    </w:r>
</w:p>
```

---

## 下一步

- [13-text-formatting.md](13-text-formatting.md) - 完整文字格式化參考
- [50-images.md](50-images.md) - 圖片與 Drawing 元素
- [64-fields.md](64-fields.md) - 欄位代碼詳解
