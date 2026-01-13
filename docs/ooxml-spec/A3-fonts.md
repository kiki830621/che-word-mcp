# 附錄 A3：字型參考

## 概述

OOXML 使用字型來定義文字的視覺外觀。字型設定涉及字型名稱、字型類型、字型替代等多個面向。

---

## 字型元素

### w:rFonts

```xml
<w:rPr>
    <w:rFonts w:ascii="Arial"
              w:hAnsi="Arial"
              w:eastAsia="標楷體"
              w:cs="Arial"/>
</w:rPr>
```

### 字型屬性

| 屬性 | 說明 | 用途 |
|------|------|------|
| `w:ascii` | ASCII 字型 | 基本拉丁字元 (0x00-0x7F) |
| `w:hAnsi` | High ANSI 字型 | 擴展拉丁字元 |
| `w:eastAsia` | 東亞字型 | 中日韓文字 |
| `w:cs` | Complex Script 字型 | 阿拉伯、希伯來等 |

### 主題字型參照

```xml
<w:rFonts w:asciiTheme="majorHAnsi"
          w:hAnsiTheme="majorHAnsi"
          w:eastAsiaTheme="majorEastAsia"
          w:cstheme="majorBidi"/>
```

| 主題字型 | 說明 |
|----------|------|
| `majorHAnsi` | 主要標題字型（西文） |
| `minorHAnsi` | 次要內文字型（西文） |
| `majorEastAsia` | 主要標題字型（東亞） |
| `minorEastAsia` | 次要內文字型（東亞） |
| `majorBidi` | 主要標題字型（雙向） |
| `minorBidi` | 次要內文字型（雙向） |

---

## 常用字型

### 西文字型

| 字型名稱 | 類型 | 用途 |
|----------|------|------|
| Times New Roman | 襯線 | 正式文件、學術 |
| Arial | 無襯線 | 通用、簡報 |
| Calibri | 無襯線 | Office 預設 |
| Cambria | 襯線 | 標題 |
| Georgia | 襯線 | 螢幕閱讀 |
| Verdana | 無襯線 | 螢幕閱讀 |
| Courier New | 等寬 | 程式碼 |
| Consolas | 等寬 | 程式碼 |

### 中文字型

| 字型名稱 | 類型 | 用途 |
|----------|------|------|
| 新細明體 | 襯線 | Windows 預設 |
| 標楷體 | 楷書 | 正式文件 |
| 微軟正黑體 | 無襯線 | 現代設計 |
| 微軟雅黑 | 無襯線 | 簡體中文 |
| 思源黑體 | 無襯線 | 開源字型 |
| 思源宋體 | 襯線 | 開源字型 |
| 蘋方 | 無襯線 | macOS |
| 華康字型系列 | 多種 | 設計用途 |

### 日文字型

| 字型名稱 | 類型 |
|----------|------|
| MS Mincho | 明朝體 |
| MS Gothic | 哥德體 |
| Meiryo | 無襯線 |
| Yu Mincho | 游明朝 |
| Yu Gothic | 游哥德 |

---

## 字型表 (fontTable.xml)

### 基本結構

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:font w:name="Calibri">
        <w:panose1 w:val="020F0502020204030204"/>
        <w:charset w:val="00"/>
        <w:family w:val="swiss"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb0="E0002AFF" w:usb1="C000247B"
               w:usb2="00000009" w:usb3="00000000"
               w:csb0="000001FF" w:csb1="00000000"/>
    </w:font>

    <w:font w:name="標楷體">
        <w:altName w:val="DFKai-SB"/>
        <w:panose1 w:val="03000509000000000000"/>
        <w:charset w:val="88"/>
        <w:family w:val="script"/>
        <w:pitch w:val="fixed"/>
    </w:font>
</w:fonts>
```

### 字型屬性說明

| 元素 | 說明 |
|------|------|
| `w:name` | 字型名稱 |
| `w:altName` | 替代名稱 |
| `w:panose1` | PANOSE 分類碼 |
| `w:charset` | 字元集代碼 |
| `w:family` | 字型家族 |
| `w:pitch` | 字元寬度 |
| `w:sig` | Unicode 簽名 |

### 字型家族 (w:family)

| 值 | 說明 |
|----|------|
| `auto` | 自動 |
| `decorative` | 裝飾性 |
| `modern` | 等寬 |
| `roman` | 襯線 |
| `script` | 手寫/楷書 |
| `swiss` | 無襯線 |

### 字元寬度 (w:pitch)

| 值 | 說明 |
|----|------|
| `default` | 預設 |
| `fixed` | 等寬 |
| `variable` | 可變寬度 |

### 字元集代碼 (w:charset)

| 值 | 說元集 |
|----|--------|
| 00 | ANSI |
| 01 | Default |
| 02 | Symbol |
| 80 | Shift-JIS (日文) |
| 81 | Hangul (韓文) |
| 86 | GB2312 (簡中) |
| 88 | Big5 (繁中) |
| A1 | Greek |
| A2 | Turkish |
| B2 | Vietnamese |
| CC | Cyrillic |
| EE | Eastern European |

---

## 主題字型定義

### theme1.xml 中的字型

```xml
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:themeElements>
        <a:fontScheme name="Office">
            <!-- 主要字型（標題） -->
            <a:majorFont>
                <a:latin typeface="Calibri Light"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="游ゴシック Light"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="等线 Light"/>
                <a:font script="Hant" typeface="新細明體"/>
            </a:majorFont>

            <!-- 次要字型（內文） -->
            <a:minorFont>
                <a:latin typeface="Calibri"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="游ゴシック"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="等线"/>
                <a:font script="Hant" typeface="新細明體"/>
            </a:minorFont>
        </a:fontScheme>
    </a:themeElements>
</a:theme>
```

### 語言腳本代碼

| script | 語言 |
|--------|------|
| `Jpan` | 日文 |
| `Hang` | 韓文 |
| `Hans` | 簡體中文 |
| `Hant` | 繁體中文 |
| `Arab` | 阿拉伯文 |
| `Hebr` | 希伯來文 |
| `Thai` | 泰文 |
| `Viet` | 越南文 |

---

## 字型預設值

### docDefaults 中的字型

```xml
<w:docDefaults>
    <w:rPrDefault>
        <w:rPr>
            <w:rFonts w:asciiTheme="minorHAnsi"
                      w:eastAsiaTheme="minorEastAsia"
                      w:hAnsiTheme="minorHAnsi"
                      w:cstheme="minorBidi"/>
            <w:sz w:val="22"/>
            <w:szCs w:val="22"/>
            <w:lang w:val="en-US" w:eastAsia="zh-TW" w:bidi="ar-SA"/>
        </w:rPr>
    </w:rPrDefault>
</w:docDefaults>
```

---

## 字型替代

### 指定替代字型

```xml
<w:font w:name="CustomFont">
    <w:altName w:val="Arial"/>
    <!-- 當 CustomFont 不存在時使用 Arial -->
</w:font>
```

### 系統字型替代順序

1. 指定的字型名稱
2. altName 替代字型
3. 同家族的系統字型
4. 系統預設字型

---

## 特殊字型

### Symbol 字型

```xml
<w:r>
    <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/>
    </w:rPr>
    <w:sym w:font="Symbol" w:char="F0B7"/>
</w:r>
```

### Wingdings 字型

```xml
<w:sym w:font="Wingdings" w:char="F0FC"/>  <!-- 打勾 -->
<w:sym w:font="Wingdings" w:char="F0FB"/>  <!-- 打叉 -->
```

### 常用符號

| 字型 | 代碼 | 符號 |
|------|------|------|
| Symbol | F0B7 | • (項目符號) |
| Symbol | F0AE | ® (註冊商標) |
| Symbol | F0D3 | © (版權) |
| Wingdings | F0FC | ✓ (打勾) |
| Wingdings | F0FB | ✗ (打叉) |
| Wingdings | F046 | ✉ (信封) |

---

## 字型嵌入

### 在 settings.xml 中啟用

```xml
<w:settings>
    <w:embedTrueTypeFonts/>
    <w:embedSystemFonts/>
    <w:saveSubsetFonts/>  <!-- 只嵌入使用的字元 -->
</w:settings>
```

### 嵌入字型檔案

```
document.docx
├── word/
│   └── fonts/
│       ├── font1.odttf  <!-- 混淆的 TrueType -->
│       └── font2.odttf
└── [Content_Types].xml
```

### Content Types

```xml
<Override PartName="/word/fonts/font1.odttf"
          ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>
```

---

## 字型相關設定

### 文件相容性設定

```xml
<w:settings>
    <!-- 使用印表機字型 -->
    <w:usePrinterMetrics/>

    <!-- 字元間距相容性 -->
    <w:doNotExpandShiftReturn/>

    <!-- 標點壓縮 -->
    <w:characterSpacingControl w:val="compressPunctuation"/>
</w:settings>
```

---

## 程式範例

### Swift

```swift
struct FontHelper {
    // 西文字型
    static let calibri = "Calibri"
    static let arial = "Arial"
    static let timesNewRoman = "Times New Roman"

    // 中文字型
    static let mingliu = "新細明體"
    static let kaiti = "標楷體"
    static let msJhengHei = "微軟正黑體"

    // 建立字型 XML
    static func fontXML(ascii: String, eastAsia: String) -> String {
        """
        <w:rFonts w:ascii="\(ascii)" w:hAnsi="\(ascii)" \
        w:eastAsia="\(eastAsia)" w:cs="\(ascii)"/>
        """
    }
}
```

### JavaScript

```javascript
const FontHelper = {
    // 西文字型
    calibri: 'Calibri',
    arial: 'Arial',
    timesNewRoman: 'Times New Roman',

    // 中文字型
    mingliu: '新細明體',
    kaiti: '標楷體',
    msJhengHei: '微軟正黑體',

    // 建立字型 XML
    fontXML(ascii, eastAsia) {
        return `<w:rFonts w:ascii="${ascii}" w:hAnsi="${ascii}" ` +
               `w:eastAsia="${eastAsia}" w:cs="${ascii}"/>`;
    }
};
```

---

## 最佳實務

### 跨平台相容性

1. 使用常見的系統字型
2. 指定合適的替代字型
3. 考慮嵌入字型（注意授權）
4. 使用主題字型以便統一管理

### 中文文件建議

- 標題：微軟正黑體 / 思源黑體
- 內文：新細明體 / 思源宋體
- 正式文件：標楷體
- 西文混排：搭配 Arial 或 Times New Roman

### 效能考量

- 減少字型種類可縮小檔案
- 使用子集嵌入減少檔案大小
- 常用字型不需嵌入

---

## 相關連結

- [文字格式](13-text-formatting.md)
- [樣式系統](30-styles.md)
- [命名空間](02-namespaces.md)
