# 表格列 (Table Row) 元素

## 概述

`w:tr` 元素代表表格中的一列，包含一或多個儲存格。

## 基本結構

```xml
<w:tr>
    <w:trPr>
        <!-- 列屬性 -->
    </w:trPr>
    <w:tc>
        <!-- 儲存格 -->
    </w:tc>
</w:tr>
```

---

## w:tr 子元素

| 元素 | 說明 | 必要 |
|------|------|------|
| `w:trPr` | 列屬性 | 否 |
| `w:tc` | 表格儲存格 | 是（至少一個） |
| `w:customXml` | 自訂 XML | 否 |
| `w:sdt` | 結構化文件標籤 | 否 |
| `w:bookmarkStart` | 書籤開始 | 否 |
| `w:bookmarkEnd` | 書籤結束 | 否 |

---

## w:trPr（列屬性）

### 完整屬性列表

| 元素 | 說明 |
|------|------|
| `w:cnfStyle` | 條件式格式 |
| `w:divId` | HTML div ID |
| `w:gridBefore` | 格線前空欄數 |
| `w:gridAfter` | 格線後空欄數 |
| `w:wBefore` | 前空寬度 |
| `w:wAfter` | 後空寬度 |
| `w:cantSplit` | 禁止跨頁分割 |
| `w:trHeight` | 列高 |
| `w:tblHeader` | 標題列 |
| `w:tblCellSpacing` | 儲存格間距 |
| `w:jc` | 列對齊 |
| `w:hidden` | 隱藏列 |
| `w:ins` | 插入修訂 |
| `w:del` | 刪除修訂 |
| `w:trPrChange` | 屬性變更修訂 |

---

## 常用列屬性詳解

### w:trHeight（列高）

```xml
<w:trPr>
    <w:trHeight w:val="720" w:hRule="exact"/>
</w:trPr>
```

| 屬性 | 說明 |
|------|------|
| `w:val` | 高度值（twips） |
| `w:hRule` | 高度規則 |

### w:hRule 高度規則

| 值 | 說明 |
|----|------|
| `auto` | 自動（根據內容） |
| `exact` | 固定高度 |
| `atLeast` | 最小高度 |

#### 範例

```xml
<!-- 固定高度 0.5 inch -->
<w:trPr>
    <w:trHeight w:val="720" w:hRule="exact"/>
</w:trPr>

<!-- 最小高度 0.25 inch -->
<w:trPr>
    <w:trHeight w:val="360" w:hRule="atLeast"/>
</w:trPr>

<!-- 自動高度 -->
<w:trPr>
    <w:trHeight w:val="0" w:hRule="auto"/>
</w:trPr>
```

### w:tblHeader（標題列）

將列標記為表格標題，在跨頁時會重複顯示。

```xml
<w:trPr>
    <w:tblHeader/>
</w:trPr>
```

**重要：**
- 標題列必須是表格的前幾列
- 可以有多個連續的標題列
- Word 會在每頁頂端重複顯示標題列

#### 多列標題範例

```xml
<w:tbl>
    <w:tblPr>...</w:tblPr>
    <w:tblGrid>...</w:tblGrid>

    <!-- 標題列 1 -->
    <w:tr>
        <w:trPr>
            <w:tblHeader/>
        </w:trPr>
        <w:tc>...</w:tc>
    </w:tr>

    <!-- 標題列 2 -->
    <w:tr>
        <w:trPr>
            <w:tblHeader/>
        </w:trPr>
        <w:tc>...</w:tc>
    </w:tr>

    <!-- 資料列（非標題） -->
    <w:tr>
        <w:tc>...</w:tc>
    </w:tr>
</w:tbl>
```

### w:cantSplit（禁止跨頁分割）

防止列被分頁符分割。

```xml
<w:trPr>
    <w:cantSplit/>
</w:trPr>
```

### w:jc（列對齊）

設定列在表格中的對齊方式。

```xml
<w:trPr>
    <w:jc w:val="center"/>
</w:trPr>
```

| 值 | 說明 |
|----|------|
| `left` | 靠左 |
| `center` | 置中 |
| `right` | 靠右 |

**注意：** 這通常用於縮排的列，而非整個表格的對齊。

### w:hidden（隱藏列）

隱藏此列。

```xml
<w:trPr>
    <w:hidden/>
</w:trPr>
```

---

## 格線偏移

### w:gridBefore / w:gridAfter（格線前後空欄）

用於建立非對稱或縮排的表格結構。

```xml
<w:trPr>
    <w:gridBefore w:val="1"/>  <!-- 跳過第一欄 -->
    <w:wBefore w:w="720" w:type="dxa"/>  <!-- 跳過的寬度 -->
</w:trPr>
```

#### 縮排表格列範例

```xml
<w:tbl>
    <w:tblGrid>
        <w:gridCol w:w="720"/>   <!-- 縮排空間 -->
        <w:gridCol w:w="2880"/>  <!-- 內容欄 1 -->
        <w:gridCol w:w="2880"/>  <!-- 內容欄 2 -->
    </w:tblGrid>

    <!-- 正常列（使用全部 3 欄） -->
    <w:tr>
        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
    </w:tr>

    <!-- 縮排列（跳過第一欄） -->
    <w:tr>
        <w:trPr>
            <w:gridBefore w:val="1"/>
            <w:wBefore w:w="720" w:type="dxa"/>
        </w:trPr>
        <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc>
    </w:tr>
</w:tbl>
```

---

## w:tblCellSpacing（儲存格間距）

覆寫表格層級的儲存格間距設定。

```xml
<w:trPr>
    <w:tblCellSpacing w:w="20" w:type="dxa"/>
</w:trPr>
```

---

## w:cnfStyle（條件式格式）

指定此列應套用的條件式格式。

```xml
<w:trPr>
    <w:cnfStyle w:val="100000000000"
                w:firstRow="1"
                w:lastRow="0"
                w:firstColumn="0"
                w:lastColumn="0"
                w:oddVBand="0"
                w:evenVBand="0"
                w:oddHBand="0"
                w:evenHBand="0"
                w:firstRowFirstColumn="0"
                w:firstRowLastColumn="0"
                w:lastRowFirstColumn="0"
                w:lastRowLastColumn="0"/>
</w:trPr>
```

### 條件式格式選項

| 屬性 | 說明 |
|------|------|
| `w:firstRow` | 首列 |
| `w:lastRow` | 末列 |
| `w:firstColumn` | 首欄 |
| `w:lastColumn` | 末欄 |
| `w:oddVBand` | 奇數垂直帶 |
| `w:evenVBand` | 偶數垂直帶 |
| `w:oddHBand` | 奇數水平帶 |
| `w:evenHBand` | 偶數水平帶 |
| `w:firstRowFirstColumn` | 首列首欄 |
| `w:firstRowLastColumn` | 首列末欄 |
| `w:lastRowFirstColumn` | 末列首欄 |
| `w:lastRowLastColumn` | 末列末欄 |

---

## 修訂追蹤

### w:ins（插入修訂）

標記此列為追蹤修訂中新增的列。

```xml
<w:trPr>
    <w:ins w:id="0" w:author="作者" w:date="2025-01-13T10:00:00Z"/>
</w:trPr>
```

### w:del（刪除修訂）

標記此列為追蹤修訂中刪除的列。

```xml
<w:trPr>
    <w:del w:id="1" w:author="作者" w:date="2025-01-13T11:00:00Z"/>
</w:trPr>
```

### w:trPrChange（屬性變更修訂）

記錄列屬性的變更。

```xml
<w:trPr>
    <w:trHeight w:val="720" w:hRule="exact"/>
    <w:trPrChange w:id="2" w:author="作者" w:date="2025-01-13T12:00:00Z">
        <w:trPr>
            <w:trHeight w:val="480" w:hRule="atLeast"/>
        </w:trPr>
    </w:trPrChange>
</w:trPr>
```

---

## 完整範例

### 標題列表格

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
        <w:trPr>
            <w:tblHeader/>
            <w:trHeight w:val="480" w:hRule="atLeast"/>
            <w:cantSplit/>
        </w:trPr>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:fill="4472C4"/>
            </w:tcPr>
            <w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r>
                    <w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>標題 A</w:t>
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
                    <w:t>標題 B</w:t>
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
                    <w:t>標題 C</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>

    <!-- 資料列 -->
    <w:tr>
        <w:trPr>
            <w:trHeight w:val="360" w:hRule="atLeast"/>
        </w:trPr>
        <w:tc><w:p><w:r><w:t>資料 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>資料 2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>資料 3</w:t></w:r></w:p></w:tc>
    </w:tr>
</w:tbl>
```

### 交替行底色（帶狀格式）

```xml
<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblLook w:firstRow="1" w:noVBand="1"/>
    </w:tblPr>
    <!-- ... -->

    <!-- 奇數列 -->
    <w:tr>
        <w:trPr>
            <w:cnfStyle w:oddHBand="1"/>
        </w:trPr>
        <!-- ... -->
    </w:tr>

    <!-- 偶數列 -->
    <w:tr>
        <w:trPr>
            <w:cnfStyle w:evenHBand="1"/>
        </w:trPr>
        <!-- ... -->
    </w:tr>
</w:tbl>
```

---

## 屬性順序

`w:trPr` 中的子元素應按以下順序排列：

1. `w:cnfStyle`
2. `w:divId`
3. `w:gridBefore`
4. `w:gridAfter`
5. `w:wBefore`
6. `w:wAfter`
7. `w:cantSplit`
8. `w:trHeight`
9. `w:tblHeader`
10. `w:tblCellSpacing`
11. `w:jc`
12. `w:hidden`
13. `w:ins`
14. `w:del`
15. `w:trPrChange`

---

## 下一步

- [22-table-cell.md](22-table-cell.md) - 表格儲存格元素
- [23-table-formatting.md](23-table-formatting.md) - 表格格式化詳解
