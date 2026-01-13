# 欄位代碼 (Field Codes)

## 概述

欄位代碼（Field Codes）是 OOXML 中的動態內容機制，用於插入可自動更新的資訊，如頁碼、日期、目錄、交互參照等。

## 欄位結構

### 基本三段式結構

```xml
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> FIELD_NAME [options] </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>顯示的結果</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### w:fldChar 類型

| fldCharType | 說明 |
|-------------|------|
| `begin` | 欄位開始 |
| `separate` | 分隔指令與結果 |
| `end` | 欄位結束 |

### 簡單欄位（無快取結果）

```xml
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> PAGE </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

---

## 常用欄位類型

### 頁碼相關

| 欄位 | 說明 | 範例結果 |
|------|------|----------|
| `PAGE` | 目前頁碼 | 5 |
| `NUMPAGES` | 總頁數 | 20 |
| `SECTIONPAGES` | 本節頁數 | 8 |
| `SECTION` | 目前節號 | 2 |

### 日期時間

| 欄位 | 說明 | 範例結果 |
|------|------|----------|
| `DATE` | 目前日期 | 2024/01/15 |
| `TIME` | 目前時間 | 10:30:00 |
| `CREATEDATE` | 建立日期 | 2024/01/01 |
| `SAVEDATE` | 儲存日期 | 2024/01/14 |
| `PRINTDATE` | 列印日期 | 2024/01/15 |

### 文件資訊

| 欄位 | 說明 |
|------|------|
| `FILENAME` | 檔案名稱 |
| `FILESIZE` | 檔案大小 |
| `AUTHOR` | 作者 |
| `TITLE` | 標題 |
| `SUBJECT` | 主題 |
| `KEYWORDS` | 關鍵字 |
| `NUMCHARS` | 字元數 |
| `NUMWORDS` | 字數 |
| `NUMPAGES` | 頁數 |

### 參照類

| 欄位 | 說明 |
|------|------|
| `REF` | 書籤參照 |
| `PAGEREF` | 書籤頁碼 |
| `NOTEREF` | 腳註/尾註參照 |
| `SEQ` | 序號 |
| `STYLEREF` | 樣式參照 |

### 連結類

| 欄位 | 說明 |
|------|------|
| `HYPERLINK` | 超連結 |
| `INCLUDEPICTURE` | 插入圖片 |
| `INCLUDETEXT` | 插入文字 |
| `LINK` | OLE 連結 |

### 目錄類

| 欄位 | 說明 |
|------|------|
| `TOC` | 目錄 |
| `TOA` | 引文目錄 |
| `INDEX` | 索引 |
| `XE` | 索引項目 |
| `TC` | 目錄項目 |

---

## 欄位格式開關

### 通用格式開關

| 開關 | 說明 | 範例 |
|------|------|------|
| `\*` | 格式開關 | `\* MERGEFORMAT` |
| `\#` | 數字格式 | `\# "0.00"` |
| `\@` | 日期格式 | `\@ "yyyy/MM/dd"` |

### 數字格式 (\#)

```xml
<w:instrText> PAGE \# "第 0 頁" </w:instrText>
<!-- 結果：第 5 頁 -->

<w:instrText> NUMPAGES \# "共 0 頁" </w:instrText>
<!-- 結果：共 20 頁 -->

<w:instrText> = 1234.5 \# "#,##0.00" </w:instrText>
<!-- 結果：1,234.50 -->
```

### 日期時間格式 (\@)

```xml
<w:instrText> DATE \@ "yyyy年M月d日" </w:instrText>
<!-- 結果：2024年1月15日 -->

<w:instrText> DATE \@ "dddd" </w:instrText>
<!-- 結果：星期一 -->

<w:instrText> TIME \@ "HH:mm:ss" </w:instrText>
<!-- 結果：10:30:45 -->

<w:instrText> CREATEDATE \@ "yyyy-MM-dd HH:mm" </w:instrText>
<!-- 結果：2024-01-01 09:00 -->
```

### 文字格式 (\*)

| 格式 | 說明 | 範例 |
|------|------|------|
| `\* Upper` | 全大寫 | HELLO |
| `\* Lower` | 全小寫 | hello |
| `\* FirstCap` | 首字大寫 | Hello world |
| `\* Caps` | 每字首字大寫 | Hello World |
| `\* MERGEFORMAT` | 保留格式 | |
| `\* CHARFORMAT` | 套用首字元格式 | |

### 數字轉文字

| 格式 | 說明 | 範例（123） |
|------|------|-------------|
| `\* Arabic` | 阿拉伯數字 | 123 |
| `\* CardText` | 基數詞 | one hundred twenty-three |
| `\* OrdText` | 序數詞 | one hundred twenty-third |
| `\* Roman` | 大寫羅馬 | CXXIII |
| `\* roman` | 小寫羅馬 | cxxiii |
| `\* Alphabetic` | 大寫字母 | D |
| `\* alphabetic` | 小寫字母 | d |
| `\* Hex` | 十六進位 | 7B |

---

## 常用欄位範例

### 頁碼

```xml
<!-- 簡單頁碼 -->
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText> PAGE </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>1</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### 「第 X 頁，共 Y 頁」

```xml
<w:p>
    <w:r>
        <w:t>第 </w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText> PAGE </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
        <w:t>1</w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r>
        <w:t> 頁，共 </w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText> NUMPAGES </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
        <w:t>10</w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r>
        <w:t> 頁</w:t>
    </w:r>
</w:p>
```

### 日期

```xml
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> DATE \@ "yyyy年MM月dd日" </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>2024年01月15日</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### 書籤參照

```xml
<!-- 定義書籤 -->
<w:p>
    <w:bookmarkStart w:id="0" w:name="ImportantSection"/>
    <w:r>
        <w:t>重要章節內容</w:t>
    </w:r>
    <w:bookmarkEnd w:id="0"/>
</w:p>

<!-- 參照書籤內容 -->
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> REF ImportantSection </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>重要章節內容</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>

<!-- 參照書籤頁碼 -->
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> PAGEREF ImportantSection </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>5</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### 序號 (SEQ)

```xml
<!-- 圖表標題 -->
<w:p>
    <w:pPr>
        <w:pStyle w:val="Caption"/>
    </w:pPr>
    <w:r>
        <w:t>圖 </w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText xml:space="preserve"> SEQ Figure \* ARABIC </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
        <w:t>1</w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r>
        <w:t>：系統架構圖</w:t>
    </w:r>
</w:p>
```

### SEQ 開關

| 開關 | 說明 |
|------|------|
| `\c` | 重複上一個序號 |
| `\h` | 隱藏結果 |
| `\n` | 插入下一個序號（預設） |
| `\r N` | 重設為 N |
| `\s 層級` | 在指定標題層級重設 |

```xml
<!-- 在標題 1 重設序號 -->
<w:instrText> SEQ Figure \* ARABIC \s 1 </w:instrText>

<!-- 重設為 10 -->
<w:instrText> SEQ Figure \r 10 </w:instrText>
```

### 超連結

```xml
<w:hyperlink r:id="rId5" w:history="1">
    <w:r>
        <w:rPr>
            <w:rStyle w:val="Hyperlink"/>
        </w:rPr>
        <w:t>點擊這裡</w:t>
    </w:r>
</w:hyperlink>
```

或使用欄位：

```xml
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> HYPERLINK "https://example.com" </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:rPr>
        <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>點擊這裡</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### HYPERLINK 開關

| 開關 | 說明 |
|------|------|
| `\l "書籤"` | 文件內書籤 |
| `\m` | 伺服器端圖片映射 |
| `\n` | 新視窗開啟 |
| `\o "提示"` | 滑鼠提示文字 |
| `\t "目標"` | 目標框架 |

```xml
<!-- 連結到文件內書籤 -->
<w:instrText> HYPERLINK \l "Chapter1" </w:instrText>

<!-- 新視窗開啟 -->
<w:instrText> HYPERLINK "https://example.com" \n </w:instrText>

<!-- 含提示文字 -->
<w:instrText> HYPERLINK "https://example.com" \o "點擊前往官網" </w:instrText>
```

---

## 計算欄位

### 表達式欄位 (=)

```xml
<!-- 簡單計算 -->
<w:instrText> = 100 + 200 </w:instrText>
<!-- 結果：300 -->

<!-- 使用書籤值 -->
<w:instrText> = Price * Quantity </w:instrText>

<!-- 使用函數 -->
<w:instrText> = ABS(-100) </w:instrText>
<!-- 結果：100 -->

<w:instrText> = ROUND(3.567, 2) </w:instrText>
<!-- 結果：3.57 -->
```

### 支援的函數

| 函數 | 說明 |
|------|------|
| `ABS(x)` | 絕對值 |
| `AND(x,y)` | 邏輯與 |
| `AVERAGE(...)` | 平均值 |
| `COUNT(...)` | 計數 |
| `DEFINED(x)` | 是否已定義 |
| `FALSE` | 假 |
| `IF(條件,真,假)` | 條件判斷 |
| `INT(x)` | 取整數 |
| `MAX(...)` | 最大值 |
| `MIN(...)` | 最小值 |
| `MOD(x,y)` | 餘數 |
| `NOT(x)` | 邏輯非 |
| `OR(x,y)` | 邏輯或 |
| `PRODUCT(...)` | 乘積 |
| `ROUND(x,n)` | 四捨五入 |
| `SIGN(x)` | 符號 |
| `SUM(...)` | 總和 |
| `TRUE` | 真 |

### 表格計算

```xml
<!-- 在表格儲存格中 -->
<w:tc>
    <w:p>
        <w:r>
            <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
            <w:instrText> =SUM(ABOVE) </w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
            <w:t>1500</w:t>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end"/>
        </w:r>
    </w:p>
</w:tc>
```

### 表格計算參照

| 參照 | 說明 |
|------|------|
| `ABOVE` | 上方儲存格 |
| `BELOW` | 下方儲存格 |
| `LEFT` | 左方儲存格 |
| `RIGHT` | 右方儲存格 |
| `A1` | 指定儲存格 |
| `A1:C3` | 儲存格範圍 |

---

## 巢狀欄位

```xml
<!-- IF 欄位巢狀 PAGE -->
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> IF </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText> PAGE </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>1</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> = 1 "首頁" "第 </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
    <w:instrText> PAGE </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>1</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> 頁" </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>首頁</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

---

## 欄位更新

### 設定自動更新

```xml
<!-- settings.xml -->
<w:settings>
    <w:updateFields w:val="true"/>
</w:settings>
```

### 鎖定欄位

```xml
<w:r>
    <w:fldChar w:fldCharType="begin" w:fldLock="true"/>
</w:r>
```

### 髒標記（需要更新）

```xml
<w:r>
    <w:fldChar w:fldCharType="begin" w:dirty="true"/>
</w:r>
```

---

## 實作注意事項

### 空格保留
- `w:instrText` 中的指令前後需要空格
- 使用 `xml:space="preserve"` 保留空格

### 快取同步
- `separate` 和 `end` 之間的內容是快取的結果
- 應用程式應在適當時機更新快取

### 格式繼承
- 使用 `\* MERGEFORMAT` 保留原始格式
- 否則結果會使用欄位代碼的格式

### 跨 Run 欄位
- 欄位可以跨多個 Run 元素
- 每個部分必須在獨立的 Run 中

---

## 相關連結

- [目錄](63-toc.md)
- [頁首頁尾](42-headers-footers.md)
- [腳註尾註](62-footnotes-endnotes.md)
- [超連結](51-hyperlinks.md)
