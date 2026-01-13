# 附錄 A4：相容性注意事項

## 概述

OOXML 文件需要在不同版本的 Word 及不同應用程式間保持相容性。本附錄說明常見的相容性問題及解決方案。

---

## Word 版本差異

### 版本與功能支援

| 功能 | Word 2007 | Word 2010 | Word 2013 | Word 2016+ |
|------|-----------|-----------|-----------|------------|
| 基本 OOXML | ✓ | ✓ | ✓ | ✓ |
| 內容控制項 | ✓ | ✓ | ✓ | ✓ |
| 核取方塊 SDT | - | ✓ | ✓ | ✓ |
| 註解回覆 | - | - | ✓ | ✓ |
| 重複區段 | - | - | ✓ | ✓ |
| 共同編輯 | - | - | - | ✓ |

### 檔案格式版本

| 副檔名 | 格式 | 說明 |
|--------|------|------|
| .docx | OOXML | Word 2007+ |
| .docm | OOXML + 巨集 | 含 VBA 巨集 |
| .dotx | OOXML 範本 | 無巨集範本 |
| .dotm | OOXML 範本 + 巨集 | 含巨集範本 |

---

## 命名空間版本

### 主要命名空間

```xml
<!-- Word 2007+ (基礎) -->
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"

<!-- Word 2010 擴展 -->
xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"

<!-- Word 2012 擴展 -->
xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"

<!-- Word 2016 擴展 -->
xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
```

### 處理未知命名空間

舊版 Word 會忽略不認識的命名空間元素，因此：
- 核心功能應使用基礎命名空間
- 進階功能可使用擴展命名空間
- 提供向後相容的替代內容

---

## 相容模式

### settings.xml 中的相容性設定

```xml
<w:settings>
    <w:compat>
        <!-- 相容性選項 -->
        <w:compatSetting w:name="compatibilityMode"
                         w:uri="http://schemas.microsoft.com/office/word"
                         w:val="15"/>  <!-- Word 2013 模式 -->

        <!-- 其他選項 -->
        <w:useFELayout/>  <!-- 使用東亞版面配置 -->
        <w:doNotExpandShiftReturn/>
        <w:adjustLineHeightInTable/>
    </w:compat>
</w:settings>
```

### 相容模式值

| val | 版本 |
|-----|------|
| 11 | Word 2003 |
| 12 | Word 2007 |
| 14 | Word 2010 |
| 15 | Word 2013+ |

---

## 常見相容性問題

### 1. 圖片定位

#### 問題
不同版本對浮動圖片定位的解釋可能不同。

#### 解決方案
```xml
<!-- 使用內嵌圖片較穩定 -->
<wp:inline>
    <!-- 圖片內容 -->
</wp:inline>

<!-- 避免複雜的錨定設定 -->
```

### 2. 表格寬度

#### 問題
表格寬度百分比在不同應用程式中可能計算不同。

#### 解決方案
```xml
<!-- 使用固定寬度較穩定 -->
<w:tblW w:w="9072" w:type="dxa"/>

<!-- 或確保百分比計算一致 -->
<w:tblW w:w="5000" w:type="pct"/>  <!-- 100% -->
```

### 3. 字型替代

#### 問題
文件使用的字型在其他系統可能不存在。

#### 解決方案
```xml
<!-- 指定替代字型 -->
<w:font w:name="CustomFont">
    <w:altName w:val="Arial"/>
</w:font>

<!-- 或使用主題字型 -->
<w:rFonts w:asciiTheme="minorHAnsi"/>
```

### 4. 數學公式

#### 問題
OMML 公式在非 Word 應用程式中可能無法正確顯示。

#### 解決方案
- 考慮轉換為圖片
- 提供 MathML 替代格式
- 使用 LaTeX 格式的備註

### 5. 巨集與 ActiveX

#### 問題
巨集和 ActiveX 控制項在其他應用程式中不支援。

#### 解決方案
- 使用 .docx（無巨集）格式
- 避免使用 ActiveX 控制項
- 使用內容控制項替代

---

## 跨應用程式相容性

### LibreOffice / OpenOffice

| 功能 | 支援程度 |
|------|----------|
| 基本文字格式 | 良好 |
| 表格 | 良好 |
| 圖片（內嵌） | 良好 |
| 圖片（浮動） | 部分 |
| 追蹤修訂 | 良好 |
| 內容控制項 | 有限 |
| OMML 公式 | 部分 |
| 樣式 | 良好 |

### Apple Pages

| 功能 | 支援程度 |
|------|----------|
| 基本文字格式 | 良好 |
| 表格 | 良好 |
| 圖片 | 良好 |
| 頁首頁尾 | 良好 |
| 追蹤修訂 | 部分 |
| 內容控制項 | 不支援 |
| 公式 | 不支援 |

### Google Docs

| 功能 | 支援程度 |
|------|----------|
| 基本文字格式 | 良好 |
| 表格 | 良好 |
| 圖片 | 良好 |
| 頁首頁尾 | 基本 |
| 追蹤修訂 | 轉換為建議 |
| 內容控制項 | 不支援 |
| 樣式 | 部分 |

---

## 最佳實務

### 建立相容性高的文件

1. **使用標準功能**
   - 避免使用過新的功能
   - 使用廣泛支援的格式

2. **字型選擇**
   - 使用常見的系統字型
   - 指定替代字型
   - 考慮嵌入字型

3. **圖片處理**
   - 優先使用內嵌圖片
   - 使用標準圖片格式（PNG, JPEG）
   - 避免複雜的文繞圖設定

4. **表格設計**
   - 使用簡單的表格結構
   - 避免複雜的合併儲存格
   - 使用固定寬度

5. **樣式使用**
   - 使用內建樣式
   - 避免過度自訂
   - 保持樣式層級簡單

### 測試建議

1. 在目標應用程式中測試開啟
2. 檢查版面配置是否正確
3. 確認字型是否正確顯示
4. 測試列印輸出

---

## Strict 與 Transitional

### 格式差異

| 特性 | Transitional | Strict |
|------|--------------|--------|
| 命名空間 | schemas.openxmlformats.org | purl.oclc.org |
| VML 支援 | 是 | 否 |
| 舊版相容 | 是 | 否 |
| ISO 標準 | ISO/IEC 29500-4 | ISO/IEC 29500-1 |

### Transitional 命名空間

```xml
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
```

### Strict 命名空間

```xml
xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
```

### 建議
- 一般文件使用 Transitional（預設）
- 需要嚴格標準合規時使用 Strict
- 大多數應用程式對 Transitional 支援較好

---

## 版本偵測

### 從 settings.xml 判斷

```xml
<!-- Word 2013+ -->
<w:compat>
    <w:compatSetting w:name="compatibilityMode"
                     w:uri="http://schemas.microsoft.com/office/word"
                     w:val="15"/>
</w:compat>
```

### 從 app.xml 判斷

```xml
<Properties>
    <Application>Microsoft Office Word</Application>
    <AppVersion>16.0000</AppVersion>  <!-- Word 2016 -->
</Properties>
```

### 程式判斷版本

```swift
func detectWordVersion(from settings: String) -> String {
    if settings.contains("compatibilityMode") {
        if settings.contains("val=\"15\"") {
            return "Word 2013+"
        } else if settings.contains("val=\"14\"") {
            return "Word 2010"
        } else if settings.contains("val=\"12\"") {
            return "Word 2007"
        }
    }
    return "Unknown"
}
```

---

## 處理舊版文件

### 開啟 .doc 文件

1. 需要額外的轉換程式
2. Binary format（非 XML）
3. 可透過 Word 轉換為 .docx

### 開啟 RTF 文件

1. 純文字格式
2. 較易解析
3. 功能有限

### 版本升級建議

```xml
<!-- 升級文件版本 -->
<w:compat>
    <w:compatSetting w:name="compatibilityMode"
                     w:uri="http://schemas.microsoft.com/office/word"
                     w:val="15"/>
</w:compat>
```

---

## 錯誤處理

### 常見錯誤

| 錯誤 | 原因 | 解決方案 |
|------|------|----------|
| 無法開啟文件 | ZIP 結構損壞 | 重新建立 ZIP |
| XML 解析錯誤 | XML 格式不正確 | 驗證 XML |
| 圖片遺失 | 關聯設定錯誤 | 檢查 .rels |
| 字型顯示錯誤 | 字型不存在 | 指定替代字型 |

### 驗證工具

1. **Office Open XML SDK**
   - 官方驗證工具
   - 詳細錯誤報告

2. **Open XML Productivity Tool**
   - 視覺化檢視
   - 比較文件

3. **線上驗證器**
   - 快速檢查
   - 基本驗證

---

## 相關連結

- [概述](01-overview.md)
- [命名空間](02-namespaces.md)
- [設定](40-section.md)
