# 表格儲存格 (Table Cell) 元素

## 概述

`w:tc` 元素代表表格中的一個儲存格，包含段落或其他區塊級內容。

## 基本結構

```xml
<w:tc>
    <w:tcPr>
        <!-- 儲存格屬性 -->
    </w:tcPr>
    <w:p>
        <!-- 段落內容 -->
    </w:p>
</w:tc>
```

**重要：** 每個儲存格必須至少包含一個 `w:p` 元素。

---

## w:tc 子元素

| 元素 | 說明 | 必要 |
|------|------|------|
| `w:tcPr` | 儲存格屬性 | 否 |
| `w:p` | 段落 | 是（至少一個） |
| `w:tbl` | 巢狀表格 | 否 |
| `w:sdt` | 結構化文件標籤 | 否 |
| `w:customXml` | 自訂 XML | 否 |
| `w:bookmarkStart` | 書籤開始 | 否 |
| `w:bookmarkEnd` | 書籤結束 | 否 |

---

## w:tcPr（儲存格屬性）

### 完整屬性列表

| 元素 | 說明 |
|------|------|
| `w:cnfStyle` | 條件式格式 |
| `w:tcW` | 儲存格寬度 |
| `w:gridSpan` | 水平合併欄數 |
| `w:hMerge` | 水平合併（舊版） |
| `w:vMerge` | 垂直合併 |
| `w:tcBorders` | 儲存格邊框 |
| `w:shd` | 儲存格底色 |
| `w:noWrap` | 不自動換行 |
| `w:tcMar` | 儲存格邊距 |
| `w:textDirection` | 文字方向 |
| `w:tcFitText` | 調整文字寬度 |
| `w:vAlign` | 垂直對齊 |
| `w:hideMark` | 隱藏標記 |
| `w:cellIns` | 插入修訂 |
| `w:cellDel` | 刪除修訂 |
| `w:cellMerge` | 合併修訂 |
| `w:tcPrChange` | 屬性變更修訂 |

---

## 常用儲存格屬性詳解

### w:tcW（儲存格寬度）

```xml
<w:tcPr>
    <w:tcW w:w="2880" w:type="dxa"/>
</w:tcPr>
```

| w:type | 說明 |
|--------|------|
| `auto` | 自動（根據內容） |
| `dxa` | Twips（1/20 pt） |
| `pct` | 百分比（5000 = 100%） |
| `nil` | 無寬度 |

#### 範例

```xml
<!-- 固定寬度 2 inch -->
<w:tcPr>
    <w:tcW w:w="2880" w:type="dxa"/>
</w:tcPr>

<!-- 百分比寬度 33% -->
<w:tcPr>
    <w:tcW w:w="1650" w:type="pct"/>
</w:tcPr>

<!-- 自動寬度 -->
<w:tcPr>
    <w:tcW w:w="0" w:type="auto"/>
</w:tcPr>
```

### w:vAlign（垂直對齊）

```xml
<w:tcPr>
    <w:vAlign w:val="center"/>
</w:tcPr>
```

| 值 | 說明 |
|----|------|
| `top` | 靠上 |
| `center` | 置中 |
| `bottom` | 靠下 |

### w:textDirection（文字方向）

```xml
<w:tcPr>
    <w:textDirection w:val="tbRl"/>
</w:tcPr>
```

| 值 | 說明 |
|----|------|
| `lrTb` | 左到右，上到下（預設） |
| `tbRl` | 上到下，右到左（直排，文字旋轉 90°） |
| `btLr` | 下到上，左到右（文字旋轉 270°） |
| `lrTbV` | 垂直：左到右 |
| `tbRlV` | 垂直：上到下 |
| `tbLrV` | 垂直：上到下，左到右 |

### w:noWrap（禁止換行）

防止儲存格內容自動換行。

```xml
<w:tcPr>
    <w:noWrap/>
</w:tcPr>
```

### w:tcFitText（調整文字寬度）

將文字壓縮以適合儲存格寬度。

```xml
<w:tcPr>
    <w:tcFitText/>
</w:tcPr>
```

---

## 儲存格合併

### w:gridSpan（水平合併）

使儲存格跨越多個格線欄。

```xml
<w:tcPr>
    <w:gridSpan w:val="3"/>  <!-- 跨越 3 欄 -->
</w:tcPr>
```

#### 水平合併範例

```xml
<w:tblGrid>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
</w:tblGrid>

<!-- 第一列：合併全部 3 欄 -->
<w:tr>
    <w:tc>
        <w:tcPr>
            <w:gridSpan w:val="3"/>
        </w:tcPr>
        <w:p><w:r><w:t>標題（跨 3 欄）</w:t></w:r></w:p>
    </w:tc>
</w:tr>

<!-- 第二列：3 個獨立儲存格 -->
<w:tr>
    <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
</w:tr>
```

### w:vMerge（垂直合併）

使儲存格與上方儲存格合併。

```xml
<!-- 合併的起始儲存格 -->
<w:tcPr>
    <w:vMerge w:val="restart"/>
</w:tcPr>

<!-- 被合併的後續儲存格 -->
<w:tcPr>
    <w:vMerge/>
</w:tcPr>
```

| w:val | 說明 |
|-------|------|
| `restart` | 開始新的垂直合併 |
| `continue`（或省略） | 繼續與上方合併 |

#### 垂直合併範例

```xml
<w:tblGrid>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
</w:tblGrid>

<!-- 列 1 -->
<w:tr>
    <w:tc>
        <w:tcPr>
            <w:vMerge w:val="restart"/>  <!-- 開始垂直合併 -->
        </w:tcPr>
        <w:p><w:r><w:t>合併儲存格</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
        <w:p><w:r><w:t>B1</w:t></w:r></w:p>
    </w:tc>
</w:tr>

<!-- 列 2 -->
<w:tr>
    <w:tc>
        <w:tcPr>
            <w:vMerge/>  <!-- 繼續垂直合併 -->
        </w:tcPr>
        <w:p/>  <!-- 必須有空段落 -->
    </w:tc>
    <w:tc>
        <w:p><w:r><w:t>B2</w:t></w:r></w:p>
    </w:tc>
</w:tr>

<!-- 列 3 -->
<w:tr>
    <w:tc>
        <w:tcPr>
            <w:vMerge/>  <!-- 繼續垂直合併 -->
        </w:tcPr>
        <w:p/>
    </w:tc>
    <w:tc>
        <w:p><w:r><w:t>B3</w:t></w:r></w:p>
    </w:tc>
</w:tr>
```

### 複合合併（水平 + 垂直）

```xml
<w:tblGrid>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
</w:tblGrid>

<!-- 列 1 -->
<w:tr>
    <w:tc>
        <w:tcPr>
            <w:gridSpan w:val="2"/>     <!-- 跨 2 欄 -->
            <w:vMerge w:val="restart"/> <!-- 開始垂直合併 -->
        </w:tcPr>
        <w:p><w:r><w:t>2×2 合併區</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
        <w:p><w:r><w:t>C1</w:t></w:r></w:p>
    </w:tc>
</w:tr>

<!-- 列 2 -->
<w:tr>
    <w:tc>
        <w:tcPr>
            <w:gridSpan w:val="2"/>
            <w:vMerge/>
        </w:tcPr>
        <w:p/>
    </w:tc>
    <w:tc>
        <w:p><w:r><w:t>C2</w:t></w:r></w:p>
    </w:tc>
</w:tr>
```

### w:hMerge（舊版水平合併）

這是舊版的水平合併方式，建議使用 `w:gridSpan` 取代。

```xml
<!-- 合併開始 -->
<w:tcPr>
    <w:hMerge w:val="restart"/>
</w:tcPr>

<!-- 繼續合併 -->
<w:tcPr>
    <w:hMerge/>
</w:tcPr>
```

---

## w:tcBorders（儲存格邊框）

覆寫表格層級的邊框設定。

```xml
<w:tcPr>
    <w:tcBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:tl2br w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:tr2bl w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tcBorders>
</w:tcPr>
```

### 邊框元素

| 元素 | 說明 |
|------|------|
| `w:top` | 上邊框 |
| `w:left` | 左邊框 |
| `w:bottom` | 下邊框 |
| `w:right` | 右邊框 |
| `w:insideH` | 內部水平邊框 |
| `w:insideV` | 內部垂直邊框 |
| `w:tl2br` | 左上到右下對角線 |
| `w:tr2bl` | 右上到左下對角線 |

### 對角線邊框範例

```xml
<w:tcPr>
    <w:tcBorders>
        <w:tl2br w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tcBorders>
</w:tcPr>
```

---

## w:shd（儲存格底色）

```xml
<w:tcPr>
    <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>
</w:tcPr>
```

### 圖案樣式

| w:val | 說明 |
|-------|------|
| `clear` | 無圖案（純色） |
| `solid` | 實心 |
| `pct5` ~ `pct95` | 百分比填充 |
| `horzStripe` | 水平條紋 |
| `vertStripe` | 垂直條紋 |
| `diagStripe` | 對角條紋 |
| `reverseDiagStripe` | 反對角條紋 |
| `horzCross` | 水平交叉 |
| `diagCross` | 對角交叉 |

### 常用底色範例

```xml
<!-- 黃色背景 -->
<w:shd w:val="clear" w:fill="FFFF00"/>

<!-- 淺灰色背景 -->
<w:shd w:val="clear" w:fill="F2F2F2"/>

<!-- 主題色彩 -->
<w:shd w:val="clear" w:themeColor="accent1" w:themeFill="accent1"/>

<!-- 50% 灰色圖案 -->
<w:shd w:val="pct50" w:color="000000" w:fill="FFFFFF"/>
```

---

## w:tcMar（儲存格邊距）

覆寫表格層級的儲存格邊距。

```xml
<w:tcPr>
    <w:tcMar>
        <w:top w:w="72" w:type="dxa"/>
        <w:left w:w="144" w:type="dxa"/>
        <w:bottom w:w="72" w:type="dxa"/>
        <w:right w:w="144" w:type="dxa"/>
    </w:tcMar>
</w:tcPr>
```

---

## w:cnfStyle（條件式格式）

指定此儲存格應套用的條件式格式。

```xml
<w:tcPr>
    <w:cnfStyle w:val="001000000000"
                w:firstRow="0"
                w:lastRow="0"
                w:firstColumn="1"
                w:lastColumn="0"
                w:oddVBand="0"
                w:evenVBand="0"
                w:oddHBand="0"
                w:evenHBand="0"
                w:firstRowFirstColumn="0"
                w:firstRowLastColumn="0"
                w:lastRowFirstColumn="0"
                w:lastRowLastColumn="0"/>
</w:tcPr>
```

---

## 修訂追蹤

### w:cellIns（插入修訂）

```xml
<w:tcPr>
    <w:cellIns w:id="0" w:author="作者" w:date="2025-01-13T10:00:00Z"/>
</w:tcPr>
```

### w:cellDel（刪除修訂）

```xml
<w:tcPr>
    <w:cellDel w:id="1" w:author="作者" w:date="2025-01-13T11:00:00Z"/>
</w:tcPr>
```

### w:cellMerge（合併修訂）

```xml
<w:tcPr>
    <w:cellMerge w:id="2" w:author="作者" w:date="2025-01-13T12:00:00Z"
                 w:vMerge="rest" w:vMergeOrig="cont"/>
</w:tcPr>
```

---

## 完整範例

### 帶格式的表格

```xml
<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="5000" w:type="pct"/>
    </w:tblPr>
    <w:tblGrid>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
    </w:tblGrid>

    <!-- 標題列 -->
    <w:tr>
        <w:trPr><w:tblHeader/></w:trPr>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="4472C4"/>
                <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>項目</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="4472C4"/>
                <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>數量</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="4472C4"/>
                <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>價格</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>

    <!-- 資料列 -->
    <w:tr>
        <w:tc>
            <w:tcPr>
                <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p><w:r><w:t>商品 A</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="right"/></w:pPr>
                <w:r><w:t>10</w:t></w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="right"/></w:pPr>
                <w:r><w:t>$100</w:t></w:r>
            </w:p>
        </w:tc>
    </w:tr>
</w:tbl>
```

### 複雜合併表格

```xml
<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="5000" w:type="pct"/>
    </w:tblPr>
    <w:tblGrid>
        <w:gridCol w:w="2000"/>
        <w:gridCol w:w="2000"/>
        <w:gridCol w:w="2000"/>
        <w:gridCol w:w="2000"/>
    </w:tblGrid>

    <!-- 列 1：大標題跨全寬 -->
    <w:tr>
        <w:tc>
            <w:tcPr>
                <w:gridSpan w:val="4"/>
                <w:shd w:val="clear" w:fill="1F4E79"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>季度報告</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>

    <!-- 列 2：分類標題 -->
    <w:tr>
        <w:tc>
            <w:tcPr>
                <w:vMerge w:val="restart"/>
                <w:shd w:val="clear" w:fill="4472C4"/>
                <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>地區</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:gridSpan w:val="3"/>
                <w:shd w:val="clear" w:fill="4472C4"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>銷售額（千元）</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>

    <!-- 列 3：子標題 -->
    <w:tr>
        <w:tc>
            <w:tcPr>
                <w:vMerge/>
            </w:tcPr>
            <w:p/>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="B4C6E7"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r><w:rPr><w:b/></w:rPr><w:t>Q1</w:t></w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="B4C6E7"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r><w:rPr><w:b/></w:rPr><w:t>Q2</w:t></w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="B4C6E7"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r><w:rPr><w:b/></w:rPr><w:t>Q3</w:t></w:r>
            </w:p>
        </w:tc>
    </w:tr>

    <!-- 資料列 -->
    <w:tr>
        <w:tc><w:p><w:r><w:t>北區</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>150</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>180</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>200</w:t></w:r></w:p></w:tc>
    </w:tr>
</w:tbl>
```

---

## 屬性順序

`w:tcPr` 中的子元素應按以下順序排列：

1. `w:cnfStyle`
2. `w:tcW`
3. `w:gridSpan`
4. `w:hMerge`
5. `w:vMerge`
6. `w:tcBorders`
7. `w:shd`
8. `w:noWrap`
9. `w:tcMar`
10. `w:textDirection`
11. `w:tcFitText`
12. `w:vAlign`
13. `w:hideMark`
14. `w:cellIns`
15. `w:cellDel`
16. `w:cellMerge`
17. `w:tcPrChange`

---

## 下一步

- [23-table-formatting.md](23-table-formatting.md) - 表格格式化詳解
- [30-styles.md](30-styles.md) - 樣式系統
