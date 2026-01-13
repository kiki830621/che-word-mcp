# 編號定義 (Numbering)

## 概述

編號定義存儲在 `word/numbering.xml` 中，控制項目符號和編號清單的外觀。

## 基本結構

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

    <!-- 抽象編號定義 -->
    <w:abstractNum w:abstractNumId="0">
        <!-- 層級定義 -->
    </w:abstractNum>

    <!-- 編號實例 -->
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"/>
    </w:num>

</w:numbering>
```

---

## w:abstractNum（抽象編號定義）

定義編號/清單的格式模板。

### 子元素

| 元素 | 說明 |
|------|------|
| `w:nsid` | 唯一識別碼 |
| `w:multiLevelType` | 多層級類型 |
| `w:tmpl` | 範本識別碼 |
| `w:name` | 名稱 |
| `w:styleLink` | 連結到樣式 |
| `w:numStyleLink` | 連結到編號樣式 |
| `w:lvl` | 層級定義 |

### w:multiLevelType 值

| 值 | 說明 |
|----|------|
| `singleLevel` | 單層級清單 |
| `multilevel` | 多層級清單（獨立） |
| `hybridMultilevel` | 混合多層級清單 |

---

## w:lvl（層級定義）

定義清單中每個層級的格式。

### 屬性

| 屬性 | 說明 |
|------|------|
| `w:ilvl` | 層級索引（0-8） |
| `w:tplc` | 範本代碼 |
| `w:tentative` | 暫定 |

### 子元素

| 元素 | 說明 |
|------|------|
| `w:start` | 起始值 |
| `w:numFmt` | 編號格式 |
| `w:lvlRestart` | 層級重啟 |
| `w:pStyle` | 段落樣式 |
| `w:isLgl` | 法律編號 |
| `w:suff` | 編號後綴 |
| `w:lvlText` | 層級文字 |
| `w:lvlPicBulletId` | 圖片項目符號 ID |
| `w:legacy` | 舊版設定 |
| `w:lvlJc` | 對齊方式 |
| `w:pPr` | 段落屬性 |
| `w:rPr` | 文字屬性 |

---

## w:numFmt（編號格式）

```xml
<w:numFmt w:val="decimal"/>
```

### 常用編號格式

| 值 | 說明 | 範例 |
|----|------|------|
| `decimal` | 十進位數字 | 1, 2, 3... |
| `upperLetter` | 大寫字母 | A, B, C... |
| `lowerLetter` | 小寫字母 | a, b, c... |
| `upperRoman` | 大寫羅馬 | I, II, III... |
| `lowerRoman` | 小寫羅馬 | i, ii, iii... |
| `bullet` | 項目符號 | •, ○, ■... |
| `ordinal` | 序數 | 1st, 2nd, 3rd... |
| `cardinalText` | 基數文字 | One, Two, Three... |
| `ordinalText` | 序數文字 | First, Second, Third... |
| `chineseCounting` | 中文計數 | 一, 二, 三... |
| `chineseLegalSimplified` | 中文大寫 | 壹, 貳, 參... |
| `taiwaneseCounting` | 台灣計數 | 一, 二, 三... |
| `taiwaneseCountingThousand` | 台灣千位 | |
| `ideographDigital` | 表意數字 | 〇, 一, 二... |
| `japaneseCounting` | 日文計數 | 一, 二, 三... |
| `aiueo` | 日文五十音 | ア, イ, ウ... |
| `iroha` | 日文伊呂波 | イ, ロ, ハ... |
| `koreanCounting` | 韓文計數 | 일, 이, 삼... |
| `koreanDigital` | 韓文數字 | 一, 二, 三... |
| `none` | 無編號 | |

---

## w:lvlText（層級文字）

定義編號的顯示格式。

```xml
<w:lvlText w:val="%1."/>  <!-- 顯示：1. 2. 3. -->
<w:lvlText w:val="(%1)"/> <!-- 顯示：(1) (2) (3) -->
<w:lvlText w:val="%1.%2."/> <!-- 多層級：1.1. 1.2. -->
```

### 佔位符

| 佔位符 | 說明 |
|--------|------|
| `%1` | 第 1 層級的值 |
| `%2` | 第 2 層級的值 |
| `%3` | 第 3 層級的值 |
| ... | ... |
| `%9` | 第 9 層級的值 |

### 範例

| w:lvlText | 輸出範例 |
|-----------|----------|
| `%1.` | 1. 2. 3. |
| `%1)` | 1) 2) 3) |
| `(%1)` | (1) (2) (3) |
| `第%1章` | 第1章 第2章 第3章 |
| `%1.%2` | 1.1 1.2 2.1 |
| `%1.%2.%3` | 1.1.1 1.1.2 1.2.1 |

---

## w:suff（編號後綴）

編號與文字之間的間隔。

```xml
<w:suff w:val="tab"/>
```

| 值 | 說明 |
|----|------|
| `tab` | 定位點（預設） |
| `space` | 空格 |
| `nothing` | 無 |

---

## w:lvlJc（編號對齊）

```xml
<w:lvlJc w:val="left"/>
```

| 值 | 說明 |
|----|------|
| `left` | 靠左 |
| `center` | 置中 |
| `right` | 靠右 |

---

## 縮排設定

在 `w:pPr` 中設定縮排。

```xml
<w:lvl w:ilvl="0">
    <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
    </w:pPr>
</w:lvl>
```

### 常用縮排值

| 層級 | w:left | w:hanging |
|------|--------|-----------|
| 0 | 720 | 360 |
| 1 | 1440 | 360 |
| 2 | 2160 | 180 |
| 3 | 2880 | 360 |

---

## 項目符號範例

### 基本項目符號清單

```xml
<w:abstractNum w:abstractNumId="0">
    <w:nsid w:val="00000001"/>
    <w:multiLevelType w:val="hybridMultilevel"/>

    <!-- 第 1 層：實心圓點 -->
    <w:lvl w:ilvl="0">
        <w:start w:val="1"/>
        <w:numFmt w:val="bullet"/>
        <w:lvlText w:val=""/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="720" w:hanging="360"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
        </w:rPr>
    </w:lvl>

    <!-- 第 2 層：空心圓 -->
    <w:lvl w:ilvl="1">
        <w:start w:val="1"/>
        <w:numFmt w:val="bullet"/>
        <w:lvlText w:val="o"/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="1440" w:hanging="360"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/>
        </w:rPr>
    </w:lvl>

    <!-- 第 3 層：實心方塊 -->
    <w:lvl w:ilvl="2">
        <w:start w:val="1"/>
        <w:numFmt w:val="bullet"/>
        <w:lvlText w:val=""/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="2160" w:hanging="360"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
        </w:rPr>
    </w:lvl>

    <!-- 層級 4-8 省略... -->
</w:abstractNum>
```

### 常用項目符號字元

| 字元 | 字型 | Unicode | 說明 |
|------|------|---------|------|
| • | Symbol | U+00B7 | 實心圓點 |
| ○ | Courier New | o | 空心圓 |
| ■ | Wingdings | U+006E | 實心方塊 |
| □ | Wingdings | U+00A8 | 空心方塊 |
| ➢ | Wingdings | U+00D8 | 箭頭 |
| ✓ | Wingdings | U+00FC | 勾選 |
| ★ | Symbol | U+00AB | 星號 |
| ◆ | Wingdings | U+0076 | 菱形 |
| ─ | Symbol | U+00BE | 橫線 |

---

## 編號清單範例

### 基本編號清單

```xml
<w:abstractNum w:abstractNumId="1">
    <w:nsid w:val="00000002"/>
    <w:multiLevelType w:val="hybridMultilevel"/>

    <!-- 第 1 層：1. 2. 3. -->
    <w:lvl w:ilvl="0">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="720" w:hanging="360"/>
        </w:pPr>
    </w:lvl>

    <!-- 第 2 層：a. b. c. -->
    <w:lvl w:ilvl="1">
        <w:start w:val="1"/>
        <w:numFmt w:val="lowerLetter"/>
        <w:lvlText w:val="%2."/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="1440" w:hanging="360"/>
        </w:pPr>
    </w:lvl>

    <!-- 第 3 層：i. ii. iii. -->
    <w:lvl w:ilvl="2">
        <w:start w:val="1"/>
        <w:numFmt w:val="lowerRoman"/>
        <w:lvlText w:val="%3."/>
        <w:lvlJc w:val="right"/>
        <w:pPr>
            <w:ind w:left="2160" w:hanging="180"/>
        </w:pPr>
    </w:lvl>
</w:abstractNum>
```

### 大綱編號（1.1, 1.2, 2.1...）

```xml
<w:abstractNum w:abstractNumId="2">
    <w:nsid w:val="00000003"/>
    <w:multiLevelType w:val="multilevel"/>

    <!-- 第 1 層：1, 2, 3 -->
    <w:lvl w:ilvl="0">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1"/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="432" w:hanging="432"/>
        </w:pPr>
    </w:lvl>

    <!-- 第 2 層：1.1, 1.2, 2.1 -->
    <w:lvl w:ilvl="1">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1.%2"/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="576" w:hanging="576"/>
        </w:pPr>
    </w:lvl>

    <!-- 第 3 層：1.1.1, 1.1.2, 1.2.1 -->
    <w:lvl w:ilvl="2">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1.%2.%3"/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="720" w:hanging="720"/>
        </w:pPr>
    </w:lvl>
</w:abstractNum>
```

### 章節編號（第一章、第二章）

```xml
<w:abstractNum w:abstractNumId="3">
    <w:nsid w:val="00000004"/>
    <w:multiLevelType w:val="multilevel"/>

    <!-- 第 1 層：第一章, 第二章 -->
    <w:lvl w:ilvl="0">
        <w:start w:val="1"/>
        <w:numFmt w:val="chineseCounting"/>
        <w:lvlText w:val="第%1章"/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="720" w:hanging="720"/>
        </w:pPr>
    </w:lvl>

    <!-- 第 2 層：第一節, 第二節 -->
    <w:lvl w:ilvl="1">
        <w:start w:val="1"/>
        <w:numFmt w:val="chineseCounting"/>
        <w:lvlText w:val="第%2節"/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
            <w:ind w:left="720" w:hanging="720"/>
        </w:pPr>
    </w:lvl>
</w:abstractNum>
```

---

## w:num（編號實例）

將抽象編號定義綁定到具體的 numId。

```xml
<w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
</w:num>

<w:num w:numId="2">
    <w:abstractNumId w:val="0"/>
    <!-- 覆寫特定層級 -->
    <w:lvlOverride w:ilvl="0">
        <w:startOverride w:val="5"/>  <!-- 從 5 開始 -->
    </w:lvlOverride>
</w:num>
```

### w:lvlOverride（層級覆寫）

```xml
<w:num w:numId="3">
    <w:abstractNumId w:val="1"/>
    <w:lvlOverride w:ilvl="0">
        <w:startOverride w:val="1"/>
        <w:lvl w:ilvl="0">
            <!-- 完全覆寫層級定義 -->
            <w:start w:val="1"/>
            <w:numFmt w:val="upperLetter"/>
            <w:lvlText w:val="%1)"/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
    </w:lvlOverride>
</w:num>
```

---

## 在段落中使用

```xml
<w:p>
    <w:pPr>
        <w:numPr>
            <w:ilvl w:val="0"/>   <!-- 層級 0 -->
            <w:numId w:val="1"/>  <!-- 使用 numId="1" -->
        </w:numPr>
    </w:pPr>
    <w:r>
        <w:t>清單項目</w:t>
    </w:r>
</w:p>
```

---

## 完整 numbering.xml 範例

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">

    <!-- 項目符號清單 -->
    <w:abstractNum w:abstractNumId="0">
        <w:nsid w:val="12345678"/>
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:tmpl w:val="87654321"/>
        <w:lvl w:ilvl="0" w:tplc="04090001">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val=""/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
            </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="1" w:tplc="04090003">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val="o"/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="1440" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>
            </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="2" w:tplc="04090005">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val=""/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="2160" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
            </w:rPr>
        </w:lvl>
        <!-- 層級 3-8 類似... -->
    </w:abstractNum>

    <!-- 編號清單 -->
    <w:abstractNum w:abstractNumId="1">
        <w:nsid w:val="ABCDEF01"/>
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:tmpl w:val="10FEDCBA"/>
        <w:lvl w:ilvl="0" w:tplc="0409000F">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="1" w:tplc="04090019">
            <w:start w:val="1"/>
            <w:numFmt w:val="lowerLetter"/>
            <w:lvlText w:val="%2."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="1440" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="2" w:tplc="0409001B">
            <w:start w:val="1"/>
            <w:numFmt w:val="lowerRoman"/>
            <w:lvlText w:val="%3."/>
            <w:lvlJc w:val="right"/>
            <w:pPr>
                <w:ind w:left="2160" w:hanging="180"/>
            </w:pPr>
        </w:lvl>
        <!-- 層級 3-8 類似... -->
    </w:abstractNum>

    <!-- 編號實例 -->
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"/>
    </w:num>
    <w:num w:numId="2">
        <w:abstractNumId w:val="1"/>
    </w:num>

</w:numbering>
```

---

## 下一步

- [40-section.md](40-section.md) - 分節屬性
- [30-styles.md](30-styles.md) - 樣式系統
