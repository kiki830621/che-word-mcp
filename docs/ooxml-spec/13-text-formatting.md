# 文字格式化屬性詳解

## 概述

`w:rPr` (Run Properties) 包含所有文字格式化設定。本文件詳細說明每個屬性的用法。

---

## 字型設定

### w:rFonts（字型）

```xml
<w:rPr>
    <w:rFonts w:ascii="Arial"           <!-- 基本拉丁文字 (U+0000-U+007F) -->
              w:hAnsi="Arial"           <!-- 高位 ANSI / 拉丁擴展 -->
              w:eastAsia="微軟正黑體"    <!-- 東亞文字 (中日韓) -->
              w:cs="Arial"/>            <!-- 複雜字型 (阿拉伯、希伯來) -->
</w:rPr>
```

#### 字型選擇屬性

| 屬性 | 說明 | 適用範圍 |
|------|------|----------|
| `w:ascii` | ASCII 字型 | U+0000-U+007F |
| `w:hAnsi` | 高位 ANSI 字型 | 拉丁擴展字元 |
| `w:eastAsia` | 東亞字型 | 中文、日文、韓文 |
| `w:cs` | 複雜字型 | 阿拉伯文、希伯來文等 |
| `w:hint` | 字型提示 | `default`, `eastAsia`, `cs` |

#### 主題字型

```xml
<w:rFonts w:asciiTheme="majorHAnsi"
          w:hAnsiTheme="majorHAnsi"
          w:eastAsiaTheme="majorEastAsia"
          w:cstheme="majorBidi"/>
```

**主題字型值：**
- `majorHAnsi`, `minorHAnsi` - 主要/次要拉丁
- `majorEastAsia`, `minorEastAsia` - 主要/次要東亞
- `majorBidi`, `minorBidi` - 主要/次要雙向

---

## 字型樣式

### w:b, w:bCs（粗體）

```xml
<!-- 啟用粗體 -->
<w:rPr>
    <w:b/>
    <w:bCs/>  <!-- 複雜字型粗體 -->
</w:rPr>

<!-- 明確啟用 -->
<w:rPr>
    <w:b w:val="true"/>
    <w:b w:val="1"/>
</w:rPr>

<!-- 停用粗體（覆寫繼承） -->
<w:rPr>
    <w:b w:val="false"/>
    <w:b w:val="0"/>
</w:rPr>
```

### w:i, w:iCs（斜體）

```xml
<w:rPr>
    <w:i/>
    <w:iCs/>  <!-- 複雜字型斜體 -->
</w:rPr>
```

### w:caps（全部大寫）

```xml
<w:rPr>
    <w:caps/>
</w:rPr>
```

顯示效果：`hello` → `HELLO`

### w:smallCaps（小型大寫）

```xml
<w:rPr>
    <w:smallCaps/>
</w:rPr>
```

顯示效果：`Hello` → `Hᴇʟʟᴏ`（小寫字母變成較小的大寫）

---

## 字型大小

### w:sz, w:szCs（字型大小）

字型大小以**半點 (half-point)** 為單位。

```xml
<w:rPr>
    <w:sz w:val="24"/>    <!-- 12pt -->
    <w:szCs w:val="24"/>  <!-- 複雜字型 12pt -->
</w:rPr>
```

**換算公式：** `w:val = 字型大小(pt) × 2`

| 字型大小 | w:val | 說明 |
|----------|-------|------|
| 8pt | 16 | |
| 9pt | 18 | |
| 10pt | 20 | |
| 10.5pt | 21 | 中文常用 |
| 11pt | 22 | |
| 12pt | 24 | 標準內文 |
| 14pt | 28 | 小標題 |
| 16pt | 32 | 標題 |
| 18pt | 36 | |
| 20pt | 40 | |
| 22pt | 44 | |
| 24pt | 48 | 大標題 |
| 26pt | 52 | |
| 28pt | 56 | |
| 36pt | 72 | |
| 48pt | 96 | |
| 72pt | 144 | |

---

## 底線

### w:u（底線）

```xml
<!-- 基本底線 -->
<w:rPr>
    <w:u w:val="single"/>
</w:rPr>

<!-- 帶顏色的底線 -->
<w:rPr>
    <w:u w:val="single" w:color="FF0000"/>
</w:rPr>

<!-- 主題色底線 -->
<w:rPr>
    <w:u w:val="single" w:themeColor="accent1"/>
</w:rPr>
```

### 底線樣式 (w:val)

| 值 | 說明 | 視覺效果 |
|----|------|----------|
| `none` | 無底線 | |
| `single` | 單底線 | ─── |
| `words` | 僅文字底線（空格無） | |
| `double` | 雙底線 | ═══ |
| `thick` | 粗底線 | ━━━ |
| `dotted` | 點線 | ··· |
| `dottedHeavy` | 粗點線 | ••• |
| `dash` | 虛線 | - - - |
| `dashedHeavy` | 粗虛線 | ━ ━ ━ |
| `dashLong` | 長虛線 | —— —— |
| `dashLongHeavy` | 粗長虛線 | |
| `dotDash` | 點虛線 | ·-·-· |
| `dashDotHeavy` | 粗點虛線 | |
| `dotDotDash` | 雙點虛線 | ··-··- |
| `dashDotDotHeavy` | 粗雙點虛線 | |
| `wave` | 波浪線 | ～～～ |
| `wavyHeavy` | 粗波浪線 | |
| `wavyDouble` | 雙波浪線 | |

---

## 刪除線

### w:strike（單刪除線）

```xml
<w:rPr>
    <w:strike/>
</w:rPr>
```

顯示效果：~~刪除文字~~

### w:dstrike（雙刪除線）

```xml
<w:rPr>
    <w:dstrike/>
</w:rPr>
```

---

## 顏色

### w:color（文字顏色）

```xml
<!-- RGB 十六進位 -->
<w:rPr>
    <w:color w:val="FF0000"/>  <!-- 紅色 -->
</w:rPr>

<!-- 自動（通常是黑色） -->
<w:rPr>
    <w:color w:val="auto"/>
</w:rPr>

<!-- 主題色彩 -->
<w:rPr>
    <w:color w:themeColor="accent1"/>
</w:rPr>

<!-- 主題色彩 + 深淺調整 -->
<w:rPr>
    <w:color w:themeColor="accent1" w:themeShade="BF"/>  <!-- 較深 -->
    <w:color w:themeColor="accent1" w:themeTint="99"/>   <!-- 較淺 -->
</w:rPr>
```

#### 主題色彩值

| 值 | 說明 |
|----|------|
| `dark1` | 深色 1（通常是黑色） |
| `light1` | 淺色 1（通常是白色） |
| `dark2` | 深色 2 |
| `light2` | 淺色 2 |
| `accent1` | 輔色 1 |
| `accent2` | 輔色 2 |
| `accent3` | 輔色 3 |
| `accent4` | 輔色 4 |
| `accent5` | 輔色 5 |
| `accent6` | 輔色 6 |
| `hyperlink` | 超連結色 |
| `followedHyperlink` | 已訪問超連結色 |

### w:highlight（螢光筆）

```xml
<w:rPr>
    <w:highlight w:val="yellow"/>
</w:rPr>
```

**可用顏色：**

| 值 | 顏色 |
|----|------|
| `black` | 黑色 |
| `blue` | 藍色 |
| `cyan` | 青色 |
| `darkBlue` | 深藍 |
| `darkCyan` | 深青 |
| `darkGray` | 深灰 |
| `darkGreen` | 深綠 |
| `darkMagenta` | 深洋紅 |
| `darkRed` | 深紅 |
| `darkYellow` | 深黃 |
| `green` | 綠色 |
| `lightGray` | 淺灰 |
| `magenta` | 洋紅 |
| `red` | 紅色 |
| `white` | 白色 |
| `yellow` | 黃色 |

### w:shd（文字底色）

比 highlight 更靈活，支援任意顏色。

```xml
<w:rPr>
    <w:shd w:val="clear" w:fill="FFFF00"/>  <!-- 黃色背景 -->
</w:rPr>

<!-- 帶圖案 -->
<w:rPr>
    <w:shd w:val="pct25" w:color="000000" w:fill="FFFFFF"/>
</w:rPr>
```

---

## 上下標

### w:vertAlign（垂直對齊）

```xml
<!-- 上標 -->
<w:rPr>
    <w:vertAlign w:val="superscript"/>
</w:rPr>

<!-- 下標 -->
<w:rPr>
    <w:vertAlign w:val="subscript"/>
</w:rPr>

<!-- 基線（預設） -->
<w:rPr>
    <w:vertAlign w:val="baseline"/>
</w:rPr>
```

#### 範例：化學式

```xml
<w:p>
    <w:r><w:t>H</w:t></w:r>
    <w:r>
        <w:rPr><w:vertAlign w:val="subscript"/></w:rPr>
        <w:t>2</w:t>
    </w:r>
    <w:r><w:t>O</w:t></w:r>
</w:p>
```

輸出：H₂O

#### 範例：數學指數

```xml
<w:p>
    <w:r><w:t>x</w:t></w:r>
    <w:r>
        <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
        <w:t>2</w:t>
    </w:r>
    <w:r><w:t> + y</w:t></w:r>
    <w:r>
        <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
        <w:t>2</w:t>
    </w:r>
    <w:r><w:t> = z</w:t></w:r>
    <w:r>
        <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
        <w:t>2</w:t>
    </w:r>
</w:p>
```

輸出：x² + y² = z²

---

## 字元間距

### w:spacing（字元間距）

以 **twips** 為單位（1 twip = 1/20 pt = 1/1440 inch）。

```xml
<!-- 擴展 2pt -->
<w:rPr>
    <w:spacing w:val="40"/>
</w:rPr>

<!-- 壓縮 1pt -->
<w:rPr>
    <w:spacing w:val="-20"/>
</w:rPr>
```

**換算：** `w:val = 間距(pt) × 20`

### w:w（字元縮放）

以百分比表示字元寬度。

```xml
<!-- 150% 寬度 -->
<w:rPr>
    <w:w w:val="150"/>
</w:rPr>

<!-- 50% 寬度（壓縮） -->
<w:rPr>
    <w:w w:val="50"/>
</w:rPr>
```

### w:kern（字距調整）

啟用字距調整，指定閾值（半點）。

```xml
<!-- 對 14pt 以上的文字啟用字距調整 -->
<w:rPr>
    <w:kern w:val="28"/>
</w:rPr>
```

### w:position（垂直位置偏移）

以**半點**為單位調整基線位置。

```xml
<!-- 上移 3pt -->
<w:rPr>
    <w:position w:val="6"/>
</w:rPr>

<!-- 下移 3pt -->
<w:rPr>
    <w:position w:val="-6"/>
</w:rPr>
```

---

## 文字效果

### w:outline（空心字）

```xml
<w:rPr>
    <w:outline/>
</w:rPr>
```

### w:shadow（陰影）

```xml
<w:rPr>
    <w:shadow/>
</w:rPr>
```

### w:emboss（浮凸）

```xml
<w:rPr>
    <w:emboss/>
</w:rPr>
```

### w:imprint（陰刻/雕刻）

```xml
<w:rPr>
    <w:imprint/>
</w:rPr>
```

### w:effect（動畫效果）

**注意：** 這是舊版 Word 的功能，現代版本不常用。

```xml
<w:rPr>
    <w:effect w:val="blinkBackground"/>
</w:rPr>
```

| 值 | 說明 |
|----|------|
| `blinkBackground` | 閃爍背景 |
| `lights` | 燈光 |
| `antsBlack` | 黑螞蟻 |
| `antsRed` | 紅螞蟻 |
| `shimmer` | 閃爍 |
| `sparkle` | 火花 |
| `none` | 無 |

---

## 文字邊框

### w:bdr（文字邊框）

```xml
<w:rPr>
    <w:bdr w:val="single"      <!-- 樣式 -->
           w:sz="4"            <!-- 寬度（1/8 pt） -->
           w:space="1"         <!-- 間距（pt） -->
           w:color="000000"    <!-- 顏色 -->
           w:frame="true"/>    <!-- 框架模式 -->
</w:rPr>
```

**邊框樣式：** 與段落邊框相同，參見 [14-paragraph-formatting.md](14-paragraph-formatting.md)

---

## 特殊屬性

### w:vanish（隱藏文字）

```xml
<w:rPr>
    <w:vanish/>
</w:rPr>
```

### w:webHidden（網頁隱藏）

只在網頁輸出時隱藏。

```xml
<w:rPr>
    <w:webHidden/>
</w:rPr>
```

### w:noProof（不校對）

跳過拼字和文法檢查。

```xml
<w:rPr>
    <w:noProof/>
</w:rPr>
```

### w:lang（語言）

```xml
<w:rPr>
    <w:lang w:val="zh-TW"        <!-- 拉丁文語言 -->
            w:eastAsia="zh-TW"   <!-- 東亞語言 -->
            w:bidi="ar-SA"/>     <!-- 雙向語言 -->
</w:rPr>
```

### w:rtl（從右到左）

```xml
<w:rPr>
    <w:rtl/>
</w:rPr>
```

### w:cs（複雜字型）

標記為複雜字型處理。

```xml
<w:rPr>
    <w:cs/>
</w:rPr>
```

---

## 強調符號

### w:em（強調符號）

東亞文字的強調標記。

```xml
<w:rPr>
    <w:em w:val="dot"/>
</w:rPr>
```

| 值 | 說明 | 效果 |
|----|------|------|
| `none` | 無 | |
| `dot` | 點 | 文字上方加點 |
| `comma` | 逗號 | 文字上方加逗號 |
| `circle` | 圓圈 | 文字上方加圓圈 |
| `underDot` | 下方點 | 文字下方加點 |

---

## 東亞版面配置

### w:eastAsianLayout

```xml
<w:rPr>
    <w:eastAsianLayout w:id="1"
                       w:combine="true"
                       w:combineBrackets="square"
                       w:vert="true"
                       w:vertCompress="true"/>
</w:rPr>
```

| 屬性 | 說明 |
|------|------|
| `w:combine` | 橫向並列（組合文字） |
| `w:combineBrackets` | 並列括號：`none`, `round`, `square`, `angle`, `curly` |
| `w:vert` | 直排 |
| `w:vertCompress` | 直排壓縮 |

---

## 完整範例

### 複雜格式文字

```xml
<w:p>
    <w:r>
        <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:eastAsia="微軟正黑體"/>
            <w:b/>
            <w:i/>
            <w:sz w:val="28"/>
            <w:color w:val="1F497D"/>
            <w:u w:val="single"/>
        </w:rPr>
        <w:t>重要標題</w:t>
    </w:r>
</w:p>
```

### 帶樣式參照

```xml
<w:r>
    <w:rPr>
        <w:rStyle w:val="Strong"/>  <!-- 參照字元樣式 -->
        <w:color w:val="FF0000"/>   <!-- 覆寫顏色 -->
    </w:rPr>
    <w:t>強調文字</w:t>
</w:r>
```

---

## 屬性順序

`w:rPr` 中的子元素應按以下順序排列（建議但非強制）：

1. `w:rStyle`
2. `w:rFonts`
3. `w:b`, `w:bCs`
4. `w:i`, `w:iCs`
5. `w:caps`, `w:smallCaps`
6. `w:strike`, `w:dstrike`
7. `w:outline`, `w:shadow`, `w:emboss`, `w:imprint`
8. `w:noProof`
9. `w:snapToGrid`
10. `w:vanish`, `w:webHidden`
11. `w:color`
12. `w:spacing`
13. `w:w`
14. `w:kern`
15. `w:position`
16. `w:sz`, `w:szCs`
17. `w:highlight`
18. `w:u`
19. `w:effect`
20. `w:bdr`
21. `w:shd`
22. `w:fitText`
23. `w:vertAlign`
24. `w:rtl`, `w:cs`
25. `w:em`
26. `w:lang`
27. `w:eastAsianLayout`
28. `w:specVanish`

---

## 下一步

- [14-paragraph-formatting.md](14-paragraph-formatting.md) - 段落格式化屬性
- [30-styles.md](30-styles.md) - 樣式系統
