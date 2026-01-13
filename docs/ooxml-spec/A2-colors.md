# 附錄 A2：色彩參考

## 概述

OOXML 支援多種色彩指定方式，包括直接 RGB 值、主題色彩、自動色彩等。

---

## 色彩指定方式

### 1. RGB 十六進位

```xml
<w:color w:val="FF0000"/>  <!-- 紅色 -->
<w:color w:val="00FF00"/>  <!-- 綠色 -->
<w:color w:val="0000FF"/>  <!-- 藍色 -->
```

格式：`RRGGBB`（不含 # 符號）

### 2. 主題色彩

```xml
<w:color w:val="4472C4" w:themeColor="accent1"/>
```

### 3. 自動色彩

```xml
<w:color w:val="auto"/>
```

自動色彩會根據背景自動選擇黑色或白色。

---

## 主題色彩

### 主題色彩名稱

| themeColor | 說明 | 預設值（Office 主題） |
|------------|------|----------------------|
| `dark1` | 深色 1（通常是黑色） | 000000 |
| `light1` | 淺色 1（通常是白色） | FFFFFF |
| `dark2` | 深色 2 | 44546A |
| `light2` | 淺色 2 | E7E6E6 |
| `accent1` | 強調色 1 | 4472C4 |
| `accent2` | 強調色 2 | ED7D31 |
| `accent3` | 強調色 3 | A5A5A5 |
| `accent4` | 強調色 4 | FFC000 |
| `accent5` | 強調色 5 | 5B9BD5 |
| `accent6` | 強調色 6 | 70AD47 |
| `hyperlink` | 超連結 | 0563C1 |
| `followedHyperlink` | 已瀏覽超連結 | 954F72 |

### 主題色彩與色調/陰影

```xml
<!-- 主題色彩加上色調（變亮） -->
<w:color w:val="8FAADC" w:themeColor="accent1" w:themeTint="99"/>

<!-- 主題色彩加上陰影（變暗） -->
<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
```

| 屬性 | 說明 | 值範圍 |
|------|------|--------|
| `w:themeTint` | 色調（混合白色） | 00-FF |
| `w:themeShade` | 陰影（混合黑色） | 00-FF |

### 計算公式

**色調計算**（混合白色）：
```
結果 = 原色 + (255 - 原色) × (themeTint / 255)
```

**陰影計算**（混合黑色）：
```
結果 = 原色 × (themeShade / 255)
```

---

## 常用色彩值

### 基本色彩

| 色彩 | RGB 值 | 名稱 |
|------|--------|------|
| 黑色 | 000000 | Black |
| 白色 | FFFFFF | White |
| 紅色 | FF0000 | Red |
| 綠色 | 00FF00 | Green / Lime |
| 藍色 | 0000FF | Blue |
| 黃色 | FFFF00 | Yellow |
| 青色 | 00FFFF | Cyan / Aqua |
| 洋紅 | FF00FF | Magenta / Fuchsia |

### 深色系列

| 色彩 | RGB 值 | 名稱 |
|------|--------|------|
| 暗紅 | 800000 | Maroon |
| 暗綠 | 008000 | Green |
| 深藍 | 000080 | Navy |
| 橄欖 | 808000 | Olive |
| 紫色 | 800080 | Purple |
| 藍綠 | 008080 | Teal |

### 灰階

| 色彩 | RGB 值 | 亮度 |
|------|--------|------|
| 黑色 | 000000 | 0% |
| 深灰 | 404040 | 25% |
| 灰色 | 808080 | 50% |
| 淺灰 | C0C0C0 | 75% |
| 銀色 | D9D9D9 | 85% |
| 白色 | FFFFFF | 100% |

### Office 2016+ 預設主題色

| 用途 | RGB 值 |
|------|--------|
| 標題深色 | 44546A |
| 內文深色 | 44546A |
| 背景淺色 | E7E6E6 |
| 強調色 1 | 4472C4 |
| 強調色 2 | ED7D31 |
| 強調色 3 | A5A5A5 |
| 強調色 4 | FFC000 |
| 強調色 5 | 5B9BD5 |
| 強調色 6 | 70AD47 |

---

## 螢光筆色彩

### 螢光筆顏色（w:highlight）

```xml
<w:rPr>
    <w:highlight w:val="yellow"/>
</w:rPr>
```

### 可用螢光筆顏色

| 值 | 色彩 | RGB 近似值 |
|----|------|-----------|
| `yellow` | 黃色 | FFFF00 |
| `green` | 綠色 | 00FF00 |
| `cyan` | 青色 | 00FFFF |
| `magenta` | 洋紅 | FF00FF |
| `blue` | 藍色 | 0000FF |
| `red` | 紅色 | FF0000 |
| `darkBlue` | 深藍 | 000080 |
| `darkCyan` | 深青 | 008080 |
| `darkGreen` | 深綠 | 008000 |
| `darkMagenta` | 深洋紅 | 800080 |
| `darkRed` | 深紅 | 800000 |
| `darkYellow` | 深黃 | 808000 |
| `darkGray` | 深灰 | 808080 |
| `lightGray` | 淺灰 | C0C0C0 |
| `black` | 黑色 | 000000 |
| `none` | 無 | - |

---

## 底紋色彩

### 段落底紋

```xml
<w:pPr>
    <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>
</w:pPr>
```

### 表格儲存格底紋

```xml
<w:tcPr>
    <w:shd w:val="clear" w:color="auto" w:fill="E7E6E6"/>
</w:tcPr>
```

### 底紋屬性

| 屬性 | 說明 |
|------|------|
| `w:val` | 圖案類型 |
| `w:color` | 圖案前景色 |
| `w:fill` | 背景填充色 |
| `w:themeFill` | 主題填充色 |
| `w:themeFillTint` | 填充色調 |
| `w:themeFillShade` | 填充陰影 |

### 底紋圖案類型 (w:val)

| 值 | 說明 |
|----|------|
| `clear` | 純色（無圖案） |
| `solid` | 實心 |
| `pct5` - `pct95` | 5%-95% 密度 |
| `horzStripe` | 水平條紋 |
| `vertStripe` | 垂直條紋 |
| `diagStripe` | 對角條紋 |
| `horzCross` | 水平交叉 |
| `diagCross` | 對角交叉 |
| `thinHorzStripe` | 細水平條紋 |
| `thinVertStripe` | 細垂直條紋 |
| `nil` | 無底紋 |

---

## 邊框色彩

### 段落邊框

```xml
<w:pBdr>
    <w:bottom w:val="single" w:sz="8" w:space="1" w:color="4472C4"/>
</w:pBdr>
```

### 表格邊框

```xml
<w:tblBorders>
    <w:top w:val="single" w:sz="4" w:color="000000"/>
    <w:bottom w:val="single" w:sz="4" w:color="000000"/>
    <w:left w:val="single" w:sz="4" w:color="000000"/>
    <w:right w:val="single" w:sz="4" w:color="000000"/>
    <w:insideH w:val="single" w:sz="4" w:color="C0C0C0"/>
    <w:insideV w:val="single" w:sz="4" w:color="C0C0C0"/>
</w:tblBorders>
```

---

## DrawingML 色彩

### srgbClr（標準 RGB）

```xml
<a:solidFill>
    <a:srgbClr val="FF0000"/>
</a:solidFill>
```

### schemeClr（主題色彩）

```xml
<a:solidFill>
    <a:schemeClr val="accent1">
        <a:lumMod val="75000"/>  <!-- 亮度 75% -->
    </a:schemeClr>
</a:solidFill>
```

### 色彩調整

| 元素 | 說明 |
|------|------|
| `a:tint` | 色調（加白） |
| `a:shade` | 陰影（加黑） |
| `a:satMod` | 飽和度調整 |
| `a:lumMod` | 亮度調整 |
| `a:alpha` | 透明度 |

### 範例：半透明藍色

```xml
<a:solidFill>
    <a:srgbClr val="0000FF">
        <a:alpha val="50000"/>  <!-- 50% 透明 -->
    </a:srgbClr>
</a:solidFill>
```

---

## 漸層

### 線性漸層

```xml
<a:gradFill>
    <a:gsLst>
        <a:gs pos="0">
            <a:srgbClr val="4472C4"/>
        </a:gs>
        <a:gs pos="100000">
            <a:srgbClr val="FFFFFF"/>
        </a:gs>
    </a:gsLst>
    <a:lin ang="5400000" scaled="1"/>  <!-- 90度（向下） -->
</a:gradFill>
```

### 漸層停止點位置

- 位置使用 1/1000 百分比
- 0 = 0%
- 50000 = 50%
- 100000 = 100%

### 角度

- 使用 1/60000 度
- 5400000 = 90度（向下）
- 0 = 0度（向右）

---

## 程式範例

### Swift

```swift
struct WordColor {
    // 解析 RGB 字串
    static func parseRGB(_ hex: String) -> (r: Int, g: Int, b: Int)? {
        guard hex.count == 6 else { return nil }
        let r = Int(hex.prefix(2), radix: 16) ?? 0
        let g = Int(hex.dropFirst(2).prefix(2), radix: 16) ?? 0
        let b = Int(hex.dropFirst(4), radix: 16) ?? 0
        return (r, g, b)
    }

    // 產生 RGB 字串
    static func toRGB(r: Int, g: Int, b: Int) -> String {
        String(format: "%02X%02X%02X", r, g, b)
    }

    // 套用色調
    static func applyTint(_ color: Int, tint: Int) -> Int {
        color + (255 - color) * tint / 255
    }

    // 套用陰影
    static func applyShade(_ color: Int, shade: Int) -> Int {
        color * shade / 255
    }
}
```

### JavaScript

```javascript
const WordColor = {
    parseRGB(hex) {
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        return { r, g, b };
    },

    toRGB(r, g, b) {
        return [r, g, b]
            .map(v => v.toString(16).padStart(2, '0'))
            .join('')
            .toUpperCase();
    },

    applyTint(color, tint) {
        return Math.round(color + (255 - color) * tint / 255);
    },

    applyShade(color, shade) {
        return Math.round(color * shade / 255);
    }
};
```

---

## 色彩無障礙

### 對比度建議

| 用途 | 最小對比度 |
|------|-----------|
| 正常文字 | 4.5:1 |
| 大型文字 | 3:1 |
| 圖形元素 | 3:1 |

### 常見高對比組合

| 前景 | 背景 | 對比度 |
|------|------|--------|
| 000000 | FFFFFF | 21:1 |
| 44546A | FFFFFF | 8.5:1 |
| 4472C4 | FFFFFF | 4.5:1 |
| FFFFFF | 4472C4 | 4.5:1 |

---

## 相關連結

- [文字格式](13-text-formatting.md)
- [段落格式](14-paragraph-formatting.md)
- [表格格式](20-table.md)
- [圖片](50-images.md)
