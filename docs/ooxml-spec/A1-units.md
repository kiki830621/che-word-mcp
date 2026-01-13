# 附錄 A1：單位換算參考

## 概述

OOXML 使用多種度量單位，不同屬性使用不同的單位。了解這些單位及其換算關係對於正確設定文件格式至關重要。

---

## 主要單位

### Twip（緹）

- **定義**：1/20 點 (point)
- **用途**：頁面大小、邊距、縮排、間距
- **換算**：
  - 1 英吋 = 1440 twips
  - 1 公分 = 567 twips
  - 1 點 = 20 twips

### Half-Point（半點）

- **定義**：1/2 點
- **用途**：字型大小 (`w:sz`, `w:szCs`)
- **換算**：
  - 12pt 字型 = 24 half-points
  - 10pt 字型 = 20 half-points

### EMU（English Metric Unit）

- **定義**：914400 EMU = 1 英吋
- **用途**：圖片尺寸、DrawingML 定位
- **換算**：
  - 1 英吋 = 914400 EMU
  - 1 公分 = 360000 EMU
  - 1 點 = 12700 EMU

### 百分比

- **表示**：通常以 50 分之一或 100 分之一表示
- **用途**：表格寬度、間距比例
- **範例**：
  - `w:w="5000"` + `w:type="pct"` = 100%（5000/50 = 100）

### Eighths of a Point（1/8 點）

- **定義**：1/8 點
- **用途**：邊框寬度 (`w:sz`)
- **換算**：
  - 1 點 = 8 eighths
  - 0.5 點 = 4 eighths

---

## 換算表

### 長度單位換算

| 從 | 到 | 公式 |
|----|-----|------|
| 英吋 | twips | × 1440 |
| 公分 | twips | × 567 |
| 毫米 | twips | × 56.7 |
| 點 | twips | × 20 |
| 英吋 | EMU | × 914400 |
| 公分 | EMU | × 360000 |
| 點 | EMU | × 12700 |
| twips | EMU | × 635 |

### 常用值對照

| 度量 | 英吋 | 公分 | Twips | EMU |
|------|------|------|-------|-----|
| 1 英吋 | 1 | 2.54 | 1440 | 914400 |
| 1 公分 | 0.394 | 1 | 567 | 360000 |
| 0.5 英吋 | 0.5 | 1.27 | 720 | 457200 |
| 1 點 | 0.0139 | 0.0353 | 20 | 12700 |

---

## 頁面尺寸

### 常用紙張大小（Twips）

| 紙張 | 寬度 | 高度 | 寬度(mm) | 高度(mm) |
|------|------|------|----------|----------|
| A4 | 11906 | 16838 | 210 | 297 |
| A3 | 16838 | 23811 | 297 | 420 |
| A5 | 8391 | 11906 | 148 | 210 |
| Letter | 12240 | 15840 | 216 | 279 |
| Legal | 12240 | 20160 | 216 | 356 |
| B5 | 10318 | 14570 | 182 | 257 |

### XML 範例

```xml
<!-- A4 直向 -->
<w:pgSz w:w="11906" w:h="16838"/>

<!-- A4 橫向 -->
<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>

<!-- Letter -->
<w:pgSz w:w="12240" w:h="15840"/>
```

---

## 邊距

### 常用邊距值（Twips）

| 描述 | 英吋 | 公分 | Twips |
|------|------|------|-------|
| 窄邊距 | 0.5" | 1.27cm | 720 |
| 標準邊距 | 1" | 2.54cm | 1440 |
| 中等邊距 | 0.75" | 1.91cm | 1080 |
| 寬邊距 | 1.25" | 3.18cm | 1800 |

### XML 範例

```xml
<!-- 標準邊距 -->
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
         w:header="720" w:footer="720" w:gutter="0"/>

<!-- 窄邊距 -->
<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"
         w:header="360" w:footer="360" w:gutter="0"/>
```

---

## 字型大小

### 常用字型大小（Half-points）

| 點數 | Half-points | 用途 |
|------|-------------|------|
| 8pt | 16 | 腳註 |
| 9pt | 18 | 小字 |
| 10pt | 20 | 內文（小） |
| 10.5pt | 21 | 內文（中文預設） |
| 11pt | 22 | 內文 |
| 12pt | 24 | 內文（大） |
| 14pt | 28 | 小標題 |
| 16pt | 32 | 標題 3 |
| 18pt | 36 | 標題 2 |
| 24pt | 48 | 標題 1 |
| 36pt | 72 | 大標題 |
| 48pt | 96 | 封面標題 |

### XML 範例

```xml
<!-- 12pt 字型 -->
<w:rPr>
    <w:sz w:val="24"/>
    <w:szCs w:val="24"/>
</w:rPr>

<!-- 10.5pt 字型 -->
<w:rPr>
    <w:sz w:val="21"/>
    <w:szCs w:val="21"/>
</w:rPr>
```

---

## 行距

### 行距單位

行距 (`w:spacing/@w:line`) 使用 twips，但解釋方式取決於 `w:lineRule`：

| lineRule | 說明 | 值的意義 |
|----------|------|----------|
| `auto` | 自動（倍數） | 值 ÷ 240 = 倍數 |
| `exact` | 精確 | twips |
| `atLeast` | 最小值 | twips |

### 常用行距

| 描述 | lineRule | 值 | 實際 |
|------|----------|-----|------|
| 單倍行距 | auto | 240 | 1.0 |
| 1.15 倍 | auto | 276 | 1.15 |
| 1.5 倍 | auto | 360 | 1.5 |
| 雙倍行距 | auto | 480 | 2.0 |
| 固定 12pt | exact | 240 | 12pt |
| 最小 14pt | atLeast | 280 | 14pt |

### XML 範例

```xml
<!-- 1.5 倍行距 -->
<w:spacing w:line="360" w:lineRule="auto"/>

<!-- 單倍行距 -->
<w:spacing w:line="240" w:lineRule="auto"/>

<!-- 固定 12pt -->
<w:spacing w:line="240" w:lineRule="exact"/>
```

---

## 縮排

### 常用縮排值（Twips）

| 描述 | 字元數 | Twips |
|------|--------|-------|
| 首行縮排 2 字 | 2 | 420（約） |
| 首行縮排 0.5" | - | 720 |
| 懸掛縮排 0.5" | - | 720 |
| 左縮排 1" | - | 1440 |

### 字元單位

使用 `w:firstLineChars` 等屬性時，單位是百分之一字元：

```xml
<!-- 首行縮排 2 個字元 -->
<w:ind w:firstLineChars="200"/>

<!-- 左縮排 3 個字元 -->
<w:ind w:leftChars="300"/>
```

---

## 表格寬度

### 寬度類型

| type | 說明 | 值的意義 |
|------|------|----------|
| `auto` | 自動 | 忽略 w 值 |
| `dxa` | twips | 固定寬度 |
| `pct` | 百分比 | 值 ÷ 50 = 百分比 |
| `nil` | 無 | 零寬度 |

### XML 範例

```xml
<!-- 100% 寬度 -->
<w:tblW w:w="5000" w:type="pct"/>

<!-- 50% 寬度 -->
<w:tblW w:w="2500" w:type="pct"/>

<!-- 固定 6 英吋 -->
<w:tblW w:w="8640" w:type="dxa"/>
```

---

## 邊框寬度

### 邊框尺寸（1/8 點）

| 描述 | 點數 | 值 |
|------|------|-----|
| 細線 | 0.5pt | 4 |
| 標準 | 1pt | 8 |
| 中等 | 1.5pt | 12 |
| 粗線 | 2.25pt | 18 |
| 很粗 | 3pt | 24 |
| 最粗 | 6pt | 48 |

### XML 範例

```xml
<!-- 1pt 邊框 -->
<w:bottom w:val="single" w:sz="8" w:color="000000"/>

<!-- 2.25pt 邊框 -->
<w:bottom w:val="single" w:sz="18" w:color="000000"/>
```

---

## 程式換算函式

### Swift 範例

```swift
struct WordUnits {
    // Twips 換算
    static func inchesToTwips(_ inches: Double) -> Int {
        Int(inches * 1440)
    }

    static func cmToTwips(_ cm: Double) -> Int {
        Int(cm * 567)
    }

    static func pointsToTwips(_ points: Double) -> Int {
        Int(points * 20)
    }

    // EMU 換算
    static func inchesToEmu(_ inches: Double) -> Int {
        Int(inches * 914400)
    }

    static func cmToEmu(_ cm: Double) -> Int {
        Int(cm * 360000)
    }

    // 字型大小
    static func pointsToHalfPoints(_ points: Double) -> Int {
        Int(points * 2)
    }

    // 行距
    static func lineSpacingMultiple(_ multiple: Double) -> Int {
        Int(multiple * 240)
    }

    // 百分比
    static func percentToFifths(_ percent: Int) -> Int {
        percent * 50
    }
}
```

### JavaScript 範例

```javascript
const WordUnits = {
    // Twips
    inchesToTwips: (inches) => Math.round(inches * 1440),
    cmToTwips: (cm) => Math.round(cm * 567),
    pointsToTwips: (points) => Math.round(points * 20),

    // EMU
    inchesToEmu: (inches) => Math.round(inches * 914400),
    cmToEmu: (cm) => Math.round(cm * 360000),

    // 字型
    pointsToHalfPoints: (points) => Math.round(points * 2),

    // 行距
    lineSpacingMultiple: (multiple) => Math.round(multiple * 240),

    // 百分比
    percentToFifths: (percent) => percent * 50
};
```

---

## 快速參考表

| 用途 | 屬性 | 單位 | 1英吋 | 1公分 |
|------|------|------|-------|-------|
| 頁面尺寸 | w:pgSz | twips | 1440 | 567 |
| 邊距 | w:pgMar | twips | 1440 | 567 |
| 縮排 | w:ind | twips | 1440 | 567 |
| 間距 | w:spacing | twips | 1440 | 567 |
| 字型大小 | w:sz | half-points | - | - |
| 圖片尺寸 | a:ext | EMU | 914400 | 360000 |
| 邊框寬度 | w:sz | 1/8 點 | - | - |
| 表格寬度% | w:w (pct) | 1/50% | - | - |

---

## 相關連結

- [頁面設定](40-section.md)
- [段落格式](14-paragraph-formatting.md)
- [文字格式](13-text-formatting.md)
- [表格結構](20-table.md)
- [圖片](50-images.md)
