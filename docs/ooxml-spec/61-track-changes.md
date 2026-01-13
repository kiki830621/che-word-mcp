# 追蹤修訂 (Track Changes)

## 概述

追蹤修訂（Track Changes）功能允許記錄文件的所有變更，包括插入、刪除、格式修改等。這些修訂記錄可以被接受或拒絕，方便多人協作編輯。

## 設定檔案

### settings.xml 中啟用追蹤修訂

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <!-- 啟用追蹤修訂 -->
    <w:trackRevisions/>

    <!-- 追蹤格式變更 -->
    <w:trackFormatting w:val="true"/>

    <!-- 追蹤移動 -->
    <w:trackMoves/>

    <!-- 修訂保護（可選） -->
    <w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>
</w:settings>
```

---

## 修訂類型

### 主要修訂元素

| 元素 | 說明 | 用途 |
|------|------|------|
| `w:ins` | 插入 | 新增的內容 |
| `w:del` | 刪除 | 被刪除的內容 |
| `w:moveFrom` | 移動來源 | 內容原本位置 |
| `w:moveTo` | 移動目標 | 內容新位置 |
| `w:rPrChange` | 字元格式變更 | Run 屬性變更 |
| `w:pPrChange` | 段落格式變更 | 段落屬性變更 |
| `w:tblPrChange` | 表格格式變更 | 表格屬性變更 |
| `w:sectPrChange` | 分節格式變更 | 分節屬性變更 |

### 修訂屬性

所有修訂元素共用這些屬性：

| 屬性 | 說明 | 範例 |
|------|------|------|
| `w:id` | 修訂唯一識別碼 | `"0"` |
| `w:author` | 修改者名稱 | `"張三"` |
| `w:date` | 修改日期時間 | `"2024-01-15T10:30:00Z"` |

---

## 插入修訂 (w:ins)

### 基本結構

```xml
<w:p>
    <w:r>
        <w:t>原有文字</w:t>
    </w:r>
    <w:ins w:id="0" w:author="張三" w:date="2024-01-15T10:30:00Z">
        <w:r>
            <w:t>新插入的文字</w:t>
        </w:r>
    </w:ins>
</w:p>
```

### 插入新段落

```xml
<w:ins w:id="1" w:author="張三" w:date="2024-01-15T10:35:00Z">
    <w:p>
        <w:r>
            <w:t>這是新插入的整個段落</w:t>
        </w:r>
    </w:p>
</w:ins>
```

---

## 刪除修訂 (w:del)

### 基本結構

```xml
<w:p>
    <w:del w:id="2" w:author="李四" w:date="2024-01-15T11:00:00Z">
        <w:r>
            <w:delText>被刪除的文字</w:delText>
        </w:r>
    </w:del>
    <w:r>
        <w:t>保留的文字</w:t>
    </w:r>
</w:p>
```

### 注意事項
- 刪除的文字使用 `w:delText` 而非 `w:t`
- 刪除的內容仍保留在文件中，只是被標記

### 刪除段落標記

```xml
<w:p>
    <w:r>
        <w:t>第一段</w:t>
    </w:r>
    <w:pPr>
        <w:rPr>
            <w:del w:id="3" w:author="李四" w:date="2024-01-15T11:05:00Z"/>
        </w:rPr>
    </w:pPr>
</w:p>
<w:p>
    <w:r>
        <w:t>原本是第二段（現在合併了）</w:t>
    </w:r>
</w:p>
```

---

## 移動修訂 (w:moveFrom / w:moveTo)

### 範例：移動段落

```xml
<!-- 原位置（來源） -->
<w:moveFrom w:id="4" w:author="王五" w:date="2024-01-15T12:00:00Z"
            w:name="move1">
    <w:p>
        <w:moveFromRangeStart w:id="5" w:name="move1"/>
        <w:r>
            <w:t>這段文字被移動了</w:t>
        </w:r>
        <w:moveFromRangeEnd w:id="5"/>
    </w:p>
</w:moveFrom>

<!-- 新位置（目標） -->
<w:moveTo w:id="6" w:author="王五" w:date="2024-01-15T12:00:00Z"
          w:name="move1">
    <w:p>
        <w:moveToRangeStart w:id="7" w:name="move1"/>
        <w:r>
            <w:t>這段文字被移動了</w:t>
        </w:r>
        <w:moveToRangeEnd w:id="7"/>
    </w:p>
</w:moveTo>
```

### 移動範圍元素

| 元素 | 說明 |
|------|------|
| `w:moveFromRangeStart` | 移動來源範圍開始 |
| `w:moveFromRangeEnd` | 移動來源範圍結束 |
| `w:moveToRangeStart` | 移動目標範圍開始 |
| `w:moveToRangeEnd` | 移動目標範圍結束 |

---

## 格式變更修訂

### 字元格式變更 (w:rPrChange)

```xml
<w:r>
    <w:rPr>
        <w:b/>  <!-- 現在是粗體 -->
        <w:rPrChange w:id="8" w:author="張三" w:date="2024-01-15T14:00:00Z">
            <w:rPr>
                <!-- 變更前不是粗體（空的 rPr） -->
            </w:rPr>
        </w:rPrChange>
    </w:rPr>
    <w:t>這段文字變成粗體了</w:t>
</w:r>
```

### 段落格式變更 (w:pPrChange)

```xml
<w:p>
    <w:pPr>
        <w:jc w:val="center"/>  <!-- 現在是置中 -->
        <w:pPrChange w:id="9" w:author="李四" w:date="2024-01-15T14:30:00Z">
            <w:pPr>
                <w:jc w:val="left"/>  <!-- 變更前是靠左 -->
            </w:pPr>
        </w:pPrChange>
    </w:pPr>
    <w:r>
        <w:t>這段變成置中了</w:t>
    </w:r>
</w:p>
```

### 樣式變更

```xml
<w:p>
    <w:pPr>
        <w:pStyle w:val="Heading1"/>  <!-- 現在是標題 1 -->
        <w:pPrChange w:id="10" w:author="王五" w:date="2024-01-15T15:00:00Z">
            <w:pPr>
                <w:pStyle w:val="Normal"/>  <!-- 變更前是內文 -->
            </w:pPr>
        </w:pPrChange>
    </w:pPr>
    <w:r>
        <w:t>這段從內文變成標題</w:t>
    </w:r>
</w:p>
```

---

## 表格修訂

### 插入列

```xml
<w:tr>
    <w:trPr>
        <w:ins w:id="11" w:author="張三" w:date="2024-01-15T16:00:00Z"/>
    </w:trPr>
    <w:tc>
        <w:p>
            <w:r>
                <w:t>新增的列</w:t>
            </w:r>
        </w:p>
    </w:tc>
</w:tr>
```

### 刪除儲存格

```xml
<w:tc>
    <w:tcPr>
        <w:cellDel w:id="12" w:author="李四" w:date="2024-01-15T16:30:00Z"/>
    </w:tcPr>
    <w:p>
        <w:del w:id="13" w:author="李四" w:date="2024-01-15T16:30:00Z">
            <w:r>
                <w:delText>被刪除的內容</w:delText>
            </w:r>
        </w:del>
    </w:p>
</w:tc>
```

### 表格屬性變更

```xml
<w:tbl>
    <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblPrChange w:id="14" w:author="王五" w:date="2024-01-15T17:00:00Z">
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
            </w:tblPr>
        </w:tblPrChange>
    </w:tblPr>
    <!-- 表格內容 -->
</w:tbl>
```

---

## 分節屬性變更

```xml
<w:sectPr>
    <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
    <w:sectPrChange w:id="15" w:author="張三" w:date="2024-01-15T18:00:00Z">
        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>  <!-- 原本是直向 -->
        </w:sectPr>
    </w:sectPrChange>
</w:sectPr>
```

---

## 編號/清單修訂

### 變更編號格式

```xml
<w:p>
    <w:pPr>
        <w:numPr>
            <w:ilvl w:val="0"/>
            <w:numId w:val="2"/>  <!-- 現在的編號 -->
            <w:numberingChange w:id="16" w:author="李四"
                               w:date="2024-01-15T19:00:00Z"
                               w:original="&lt;w:numId w:val=&quot;1&quot;/&gt;"/>
        </w:numPr>
    </w:pPr>
    <w:r>
        <w:t>清單項目</w:t>
    </w:r>
</w:p>
```

---

## 完整範例

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <!-- 原有段落 -->
        <w:p>
            <w:r>
                <w:t>這是原有的文字，</w:t>
            </w:r>
            <!-- 刪除的內容 -->
            <w:del w:id="0" w:author="審稿人" w:date="2024-01-15T10:00:00Z">
                <w:r>
                    <w:delText>多餘的</w:delText>
                </w:r>
            </w:del>
            <w:r>
                <w:t>這裡</w:t>
            </w:r>
            <!-- 插入的內容 -->
            <w:ins w:id="1" w:author="作者" w:date="2024-01-15T11:00:00Z">
                <w:r>
                    <w:t>新增的說明</w:t>
                </w:r>
            </w:ins>
            <w:r>
                <w:t>。</w:t>
            </w:r>
        </w:p>

        <!-- 格式變更的段落 -->
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
                <w:pPrChange w:id="2" w:author="編輯" w:date="2024-01-15T12:00:00Z">
                    <w:pPr>
                        <w:jc w:val="left"/>
                    </w:pPr>
                </w:pPrChange>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:rPrChange w:id="3" w:author="編輯" w:date="2024-01-15T12:00:00Z">
                        <w:rPr/>
                    </w:rPrChange>
                </w:rPr>
                <w:t>這段文字被設為粗體和置中</w:t>
            </w:r>
        </w:p>

        <!-- 新插入的段落 -->
        <w:ins w:id="4" w:author="作者" w:date="2024-01-15T13:00:00Z">
            <w:p>
                <w:r>
                    <w:t>這是完全新增的段落。</w:t>
                </w:r>
            </w:p>
        </w:ins>

        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>
        </w:sectPr>
    </w:body>
</w:document>
```

---

## 接受/拒絕修訂

### 接受插入
移除 `w:ins` 標籤，保留內部內容

### 接受刪除
移除整個 `w:del` 區塊

### 拒絕插入
移除整個 `w:ins` 區塊

### 拒絕刪除
移除 `w:del` 標籤，將 `w:delText` 改為 `w:t`

### 接受格式變更
移除 `w:*PrChange` 元素，保留當前格式

### 拒絕格式變更
用 `w:*PrChange` 內的舊格式取代當前格式

---

## 實作注意事項

### ID 管理
- 所有修訂的 ID 必須在文件中唯一
- ID 用於追蹤和配對相關修訂

### 作者資訊
- 作者名稱通常來自應用程式設定
- 可以透過 settings.xml 設定預設作者

### 日期格式
- 使用 ISO 8601 格式
- 建議使用 UTC 時間

### 嵌套規則
- 修訂可以嵌套（如插入後再修改格式）
- 移動修訂需要配對的 moveFrom 和 moveTo

---

## 相關連結

- [註解](60-comments.md)
- [段落結構](11-paragraph.md)
- [文字格式](13-text-formatting.md)
- [表格結構](20-table.md)
