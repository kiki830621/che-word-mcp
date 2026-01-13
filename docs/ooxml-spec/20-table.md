# 表格 (Table) 結構

## 概述

`w:tbl` 是 WordprocessingML 中的表格元素，包含完整的表格結構和格式設定。

## 基本結構

```xml
<w:tbl>
    <w:tblPr>
        <!-- 表格屬性 -->
    </w:tblPr>
    <w:tblGrid>
        <!-- 欄寬定義 -->
    </w:tblGrid>
    <w:tr>
        <!-- 表格列 -->
    </w:tr>
</w:tbl>
```

---

## w:tbl 子元素

| 元素 | 說明 | 必要 |
|------|------|------|
| `w:tblPr` | 表格屬性 | 是 |
| `w:tblGrid` | 欄格線定義 | 是 |
| `w:tr` | 表格列 | 是（至少一個） |
| `w:bookmarkStart` | 書籤開始 | 否 |
| `w:bookmarkEnd` | 書籤結束 | 否 |
| `w:customXml` | 自訂 XML | 否 |
| `w:sdt` | 結構化文件標籤 | 否 |

---

## w:tblPr（表格屬性）

### 完整屬性列表

| 元素 | 說明 |
|------|------|
| `w:tblStyle` | 表格樣式 |
| `w:tblpPr` | 表格定位 |
| `w:tblOverlap` | 重疊設定 |
| `w:bidiVisual` | 雙向視覺排列 |
| `w:tblStyleRowBandSize` | 樣式列帶大小 |
| `w:tblStyleColBandSize` | 樣式欄帶大小 |
| `w:tblW` | 表格寬度 |
| `w:jc` | 表格對齊 |
| `w:tblCellSpacing` | 儲存格間距 |
| `w:tblInd` | 表格縮排 |
| `w:tblBorders` | 表格邊框 |
| `w:shd` | 表格底色 |
| `w:tblLayout` | 表格配置 |
| `w:tblCellMar` | 預設儲存格邊距 |
| `w:tblLook` | 表格外觀選項 |
| `w:tblCaption` | 表格標題（輔助功能） |
| `w:tblDescription` | 表格描述（輔助功能） |

---

## 常用表格屬性詳解

### w:tblStyle（表格樣式）

```xml
<w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
</w:tblPr>
```

#### 常用內建樣式

| 樣式 ID | 說明 |
|---------|------|
| `TableGrid` | 格線表格 |
| `TableNormal` | 標準表格 |
| `LightShading` | 淺色網底 |
| `LightList` | 淺色清單 |
| `LightGrid` | 淺色格線 |
| `MediumShading1` | 中等網底 1 |
| `MediumShading2` | 中等網底 2 |
| `MediumList1` | 中等清單 1 |
| `MediumList2` | 中等清單 2 |
| `MediumGrid1` | 中等格線 1 |
| `MediumGrid2` | 中等格線 2 |
| `MediumGrid3` | 中等格線 3 |
| `DarkList` | 深色清單 |
| `ColorfulShading` | 彩色網底 |
| `ColorfulList` | 彩色清單 |
| `ColorfulGrid` | 彩色格線 |

### w:tblW（表格寬度）

```xml
<!-- 百分比寬度（5000 = 100%） -->
<w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/>
</w:tblPr>

<!-- 固定寬度（twips） -->
<w:tblPr>
    <w:tblW w:w="9000" w:type="dxa"/>
</w:tblPr>

<!-- 自動寬度 -->
<w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
</w:tblPr>
```

| w:type | 說明 |
|--------|------|
| `auto` | 自動（根據內容） |
| `dxa` | Twips（1/20 pt） |
| `pct` | 百分比（5000 = 100%） |
| `nil` | 無寬度 |

### w:jc（表格對齊）

```xml
<w:tblPr>
    <w:jc w:val="center"/>
</w:tblPr>
```

| 值 | 說明 |
|----|------|
| `left` | 靠左 |
| `center` | 置中 |
| `right` | 靠右 |

### w:tblInd（表格縮排）

```xml
<w:tblPr>
    <w:tblInd w:w="720" w:type="dxa"/>
</w:tblPr>
```

### w:tblLayout（表格配置）

```xml
<w:tblPr>
    <w:tblLayout w:type="fixed"/>
</w:tblPr>
```

| 值 | 說明 |
|----|------|
| `autofit` | 自動調整（根據內容） |
| `fixed` | 固定（不隨內容調整） |

---

## w:tblBorders（表格邊框）

```xml
<w:tblPr>
    <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
</w:tblPr>
```

### 邊框元素

| 元素 | 說明 |
|------|------|
| `w:top` | 表格上邊框 |
| `w:left` | 表格左邊框 |
| `w:bottom` | 表格下邊框 |
| `w:right` | 表格右邊框 |
| `w:insideH` | 內部水平邊框 |
| `w:insideV` | 內部垂直邊框 |

### 邊框屬性

| 屬性 | 說明 |
|------|------|
| `w:val` | 邊框樣式（參見 [14-paragraph-formatting.md](14-paragraph-formatting.md)） |
| `w:sz` | 邊框寬度（1/8 pt） |
| `w:space` | 與內容間距（pt） |
| `w:color` | 顏色（RGB hex） |
| `w:themeColor` | 主題色彩 |

### 無邊框表格

```xml
<w:tblPr>
    <w:tblBorders>
        <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:tblBorders>
</w:tblPr>
```

---

## w:tblCellMar（預設儲存格邊距）

```xml
<w:tblPr>
    <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>
</w:tblPr>
```

### 邊距元素

| 元素 | 說明 | 預設值 |
|------|------|--------|
| `w:top` | 上邊距 | 0 |
| `w:left` | 左邊距 | 108 twips |
| `w:bottom` | 下邊距 | 0 |
| `w:right` | 右邊距 | 108 twips |

---

## w:tblCellSpacing（儲存格間距）

```xml
<w:tblPr>
    <w:tblCellSpacing w:w="20" w:type="dxa"/>
</w:tblPr>
```

在儲存格之間添加間距（類似 HTML 的 cellspacing）。

---

## w:tblLook（表格外觀選項）

控制表格樣式的條件格式應用。

```xml
<w:tblPr>
    <w:tblLook w:val="04A0"
               w:firstRow="1"
               w:lastRow="0"
               w:firstColumn="1"
               w:lastColumn="0"
               w:noHBand="0"
               w:noVBand="1"/>
</w:tblPr>
```

### 外觀選項

| 屬性 | 說明 |
|------|------|
| `w:firstRow` | 啟用首列格式 |
| `w:lastRow` | 啟用末列格式 |
| `w:firstColumn` | 啟用首欄格式 |
| `w:lastColumn` | 啟用末欄格式 |
| `w:noHBand` | 停用水平帶狀格式 |
| `w:noVBand` | 停用垂直帶狀格式 |

### w:val 位元遮罩

`w:val` 是一個十六進位值，編碼所有選項：

| 位元 | 說明 |
|------|------|
| 0x0020 | 首列 |
| 0x0040 | 末列 |
| 0x0080 | 首欄 |
| 0x0100 | 末欄 |
| 0x0200 | 無水平帶 |
| 0x0400 | 無垂直帶 |

---

## w:tblGrid（欄格線）

定義表格的欄結構。

```xml
<w:tblGrid>
    <w:gridCol w:w="2880"/>  <!-- 欄 1 寬度 -->
    <w:gridCol w:w="2880"/>  <!-- 欄 2 寬度 -->
    <w:gridCol w:w="2880"/>  <!-- 欄 3 寬度 -->
</w:tblGrid>
```

### w:gridCol 屬性

| 屬性 | 說明 |
|------|------|
| `w:w` | 欄寬（twips） |

**注意：**
- `w:tblGrid` 中的 `w:gridCol` 數量決定表格的欄數
- 欄寬可以被儲存格覆寫
- 合併儲存格時，儲存格跨越多個格線欄

---

## w:tblpPr（浮動表格定位）

將表格設為浮動定位。

```xml
<w:tblPr>
    <w:tblpPr w:leftFromText="180"
              w:rightFromText="180"
              w:topFromText="0"
              w:bottomFromText="0"
              w:vertAnchor="text"
              w:horzAnchor="margin"
              w:tblpXSpec="center"
              w:tblpY="720"/>
</w:tblPr>
```

### 定位屬性

| 屬性 | 說明 |
|------|------|
| `w:leftFromText` | 與左側文字距離 (twips) |
| `w:rightFromText` | 與右側文字距離 (twips) |
| `w:topFromText` | 與上方文字距離 (twips) |
| `w:bottomFromText` | 與下方文字距離 (twips) |
| `w:vertAnchor` | 垂直錨點：`text`, `margin`, `page` |
| `w:horzAnchor` | 水平錨點：`text`, `margin`, `page` |
| `w:tblpX` | 水平位置偏移 (twips) |
| `w:tblpY` | 垂直位置偏移 (twips) |
| `w:tblpXSpec` | 水平對齊：`left`, `center`, `right`, `inside`, `outside` |
| `w:tblpYSpec` | 垂直對齊：`top`, `center`, `bottom`, `inside`, `outside` |

---

## 完整表格範例

### 基本三欄表格

```xml
<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>
    </w:tblPr>
    <w:tblGrid>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
    </w:tblGrid>

    <!-- 標題列 -->
    <w:tr>
        <w:trPr>
            <w:tblHeader/>
        </w:trPr>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="4472C4"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>欄位 A</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="4472C4"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>欄位 B</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="4472C4"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>欄位 C</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>

    <!-- 資料列 1 -->
    <w:tr>
        <w:tc>
            <w:p><w:r><w:t>資料 1</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
            <w:p><w:r><w:t>資料 2</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
            <w:p><w:r><w:t>資料 3</w:t></w:r></w:p>
        </w:tc>
    </w:tr>

    <!-- 資料列 2 -->
    <w:tr>
        <w:tc>
            <w:p><w:r><w:t>資料 4</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
            <w:p><w:r><w:t>資料 5</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
            <w:p><w:r><w:t>資料 6</w:t></w:r></w:p>
        </w:tc>
    </w:tr>
</w:tbl>
```

### 置中表格

```xml
<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="0" w:type="auto"/>
        <w:jc w:val="center"/>
    </w:tblPr>
    <!-- ... -->
</w:tbl>
```

### 無邊框表格（用於版面配置）

```xml
<w:tbl>
    <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
            <w:top w:val="none"/>
            <w:left w:val="none"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
            <w:insideH w:val="none"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
        <w:tblCellMar>
            <w:top w:w="0" w:type="dxa"/>
            <w:left w:w="0" w:type="dxa"/>
            <w:bottom w:w="0" w:type="dxa"/>
            <w:right w:w="0" w:type="dxa"/>
        </w:tblCellMar>
    </w:tblPr>
    <!-- ... -->
</w:tbl>
```

---

## 下一步

- [21-table-row.md](21-table-row.md) - 表格列元素
- [22-table-cell.md](22-table-cell.md) - 表格儲存格元素
- [23-table-formatting.md](23-table-formatting.md) - 表格格式化詳解
