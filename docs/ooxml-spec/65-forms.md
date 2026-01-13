# 表單控制項 (Form Controls)

## 概述

OOXML 支援多種表單控制項，讓使用者可以在文件中填寫資料。主要有兩種類型：
1. **傳統表單欄位**（Legacy Form Fields）- 舊式但相容性高
2. **內容控制項**（Content Controls / SDT）- 新式功能更豐富

---

## 內容控制項 (SDT - Structured Document Tags)

### 基本結構

```xml
<w:sdt>
    <w:sdtPr>
        <!-- 控制項屬性 -->
    </w:sdtPr>
    <w:sdtEndPr>
        <!-- 結束屬性（可選） -->
    </w:sdtEndPr>
    <w:sdtContent>
        <!-- 控制項內容 -->
    </w:sdtContent>
</w:sdt>
```

### 通用 SDT 屬性

```xml
<w:sdtPr>
    <!-- 標籤（程式識別用） -->
    <w:tag w:val="field_name"/>

    <!-- ID（唯一識別碼） -->
    <w:id w:val="12345678"/>

    <!-- 別名（顯示名稱） -->
    <w:alias w:val="姓名欄位"/>

    <!-- 鎖定設定 -->
    <w:lock w:val="sdtLocked"/>  <!-- 不能刪除 SDT -->

    <!-- 提示文字 -->
    <w:placeholder>
        <w:docPart w:val="DefaultPlaceholder_Text"/>
    </w:placeholder>

    <!-- 顯示為方塊或標籤 -->
    <w:showingPlcHdr/>  <!-- 顯示預留位置 -->

    <!-- 暫時性（填寫後移除標籤） -->
    <w:temporary/>
</w:sdtPr>
```

### 鎖定選項 (w:lock)

| 值 | 說明 |
|----|------|
| `contentLocked` | 內容不可編輯 |
| `sdtContentLocked` | SDT 和內容都不可編輯 |
| `sdtLocked` | SDT 不可刪除 |
| （無） | 無限制 |

---

## 純文字控制項

### 結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="name"/>
        <w:id w:val="1"/>
        <w:alias w:val="姓名"/>
        <w:text/>  <!-- 標記為純文字 -->
    </w:sdtPr>
    <w:sdtContent>
        <w:r>
            <w:t>請輸入姓名</w:t>
        </w:r>
    </w:sdtContent>
</w:sdt>
```

### 多行文字

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="description"/>
        <w:id w:val="2"/>
        <w:text w:multiLine="1"/>  <!-- 允許多行 -->
    </w:sdtPr>
    <w:sdtContent>
        <w:p>
            <w:r>
                <w:t>第一行</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>第二行</w:t>
            </w:r>
        </w:p>
    </w:sdtContent>
</w:sdt>
```

---

## Rich Text 控制項

### 結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="richtext_field"/>
        <w:id w:val="3"/>
        <!-- 不加 <w:text/> 即為 Rich Text -->
    </w:sdtPr>
    <w:sdtContent>
        <w:p>
            <w:r>
                <w:rPr>
                    <w:b/>
                </w:rPr>
                <w:t>粗體文字</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:i/>
                </w:rPr>
                <w:t>斜體文字</w:t>
            </w:r>
        </w:p>
    </w:sdtContent>
</w:sdt>
```

---

## 下拉式清單

### 結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="department"/>
        <w:id w:val="4"/>
        <w:alias w:val="部門"/>
        <w:dropDownList>
            <w:listItem w:displayText="請選擇..." w:value=""/>
            <w:listItem w:displayText="業務部" w:value="sales"/>
            <w:listItem w:displayText="工程部" w:value="engineering"/>
            <w:listItem w:displayText="人資部" w:value="hr"/>
            <w:listItem w:displayText="財務部" w:value="finance"/>
        </w:dropDownList>
    </w:sdtPr>
    <w:sdtContent>
        <w:r>
            <w:t>請選擇...</w:t>
        </w:r>
    </w:sdtContent>
</w:sdt>
```

### w:listItem 屬性

| 屬性 | 說明 |
|------|------|
| `w:displayText` | 顯示文字 |
| `w:value` | 實際值 |

---

## 下拉式方塊（可編輯）

### 結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="country"/>
        <w:id w:val="5"/>
        <w:comboBox>  <!-- 可編輯的下拉 -->
            <w:listItem w:displayText="台灣" w:value="TW"/>
            <w:listItem w:displayText="日本" w:value="JP"/>
            <w:listItem w:displayText="美國" w:value="US"/>
        </w:comboBox>
    </w:sdtPr>
    <w:sdtContent>
        <w:r>
            <w:t>台灣</w:t>
        </w:r>
    </w:sdtContent>
</w:sdt>
```

---

## 日期選擇器

### 結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="birthdate"/>
        <w:id w:val="6"/>
        <w:alias w:val="出生日期"/>
        <w:date w:fullDate="2024-01-15T00:00:00Z">
            <w:dateFormat w:val="yyyy/MM/dd"/>
            <w:lid w:val="zh-TW"/>
            <w:storeMappedDataAs w:val="dateTime"/>
            <w:calendar w:val="gregorian"/>
        </w:date>
    </w:sdtPr>
    <w:sdtContent>
        <w:r>
            <w:t>2024/01/15</w:t>
        </w:r>
    </w:sdtContent>
</w:sdt>
```

### 日期屬性

| 元素 | 說明 |
|------|------|
| `w:fullDate` | 完整日期值（ISO 格式） |
| `w:dateFormat` | 顯示格式 |
| `w:lid` | 語言代碼 |
| `w:storeMappedDataAs` | 儲存格式 |
| `w:calendar` | 曆法類型 |

### 曆法類型 (w:calendar)

| 值 | 說明 |
|----|------|
| `gregorian` | 西曆 |
| `taiwan` | 民國曆 |
| `japan` | 日本年號 |
| `hijri` | 伊斯蘭曆 |
| `hebrew` | 希伯來曆 |

---

## 核取方塊

### 結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="agree"/>
        <w:id w:val="7"/>
        <w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
            <w14:checked w14:val="0"/>  <!-- 0=未勾選, 1=已勾選 -->
            <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>  <!-- ☒ -->
            <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>  <!-- ☐ -->
        </w14:checkbox>
    </w:sdtPr>
    <w:sdtContent>
        <w:r>
            <w:rPr>
                <w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/>
            </w:rPr>
            <w:t>☐</w:t>
        </w:r>
    </w:sdtContent>
</w:sdt>
```

### 常用核取方塊符號

| Unicode | 字元 | 說明 |
|---------|------|------|
| 2610 | ☐ | 未勾選方塊 |
| 2611 | ☑ | 勾選方塊 |
| 2612 | ☒ | 叉叉方塊 |

---

## 圖片控制項

### 結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="photo"/>
        <w:id w:val="8"/>
        <w:picture/>
    </w:sdtPr>
    <w:sdtContent>
        <w:r>
            <w:drawing>
                <!-- DrawingML 圖片 -->
            </w:drawing>
        </w:r>
    </w:sdtContent>
</w:sdt>
```

---

## 重複區段控制項

### 結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:tag w:val="items"/>
        <w:id w:val="9"/>
        <w15:repeatingSection xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>
    </w:sdtPr>
    <w:sdtContent>
        <!-- 重複區段項目 -->
        <w:sdt>
            <w:sdtPr>
                <w15:repeatingSectionItem xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r>
                        <w:t>項目 1</w:t>
                    </w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>
        <!-- 更多項目 -->
    </w:sdtContent>
</w:sdt>
```

---

## 傳統表單欄位

### 文字欄位

```xml
<w:r>
    <w:fldChar w:fldCharType="begin">
        <w:ffData>
            <w:name w:val="Text1"/>
            <w:enabled/>
            <w:calcOnExit w:val="0"/>
            <w:textInput>
                <w:type w:val="regular"/>
                <w:maxLength w:val="100"/>
                <w:default w:val="預設值"/>
            </w:textInput>
        </w:ffData>
    </w:fldChar>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> FORMTEXT </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
    <w:t>預設值</w:t>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### 文字類型 (w:type)

| 值 | 說明 |
|----|------|
| `regular` | 一般文字 |
| `number` | 數字 |
| `date` | 日期 |
| `currentDate` | 目前日期 |
| `currentTime` | 目前時間 |
| `calculated` | 計算欄位 |

### 核取方塊（傳統）

```xml
<w:r>
    <w:fldChar w:fldCharType="begin">
        <w:ffData>
            <w:name w:val="Check1"/>
            <w:enabled/>
            <w:calcOnExit w:val="0"/>
            <w:checkBox>
                <w:sizeAuto/>
                <w:default w:val="0"/>  <!-- 0=未勾, 1=已勾 -->
            </w:checkBox>
        </w:ffData>
    </w:fldChar>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

### 下拉選單（傳統）

```xml
<w:r>
    <w:fldChar w:fldCharType="begin">
        <w:ffData>
            <w:name w:val="Dropdown1"/>
            <w:enabled/>
            <w:calcOnExit w:val="0"/>
            <w:ddList>
                <w:result w:val="0"/>  <!-- 選取的索引 -->
                <w:listEntry w:val="選項 A"/>
                <w:listEntry w:val="選項 B"/>
                <w:listEntry w:val="選項 C"/>
            </w:ddList>
        </w:ffData>
    </w:fldChar>
</w:r>
<w:r>
    <w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText>
</w:r>
<w:r>
    <w:fldChar w:fldCharType="end"/>
</w:r>
```

---

## 完整表單範例

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
    <w:body>
        <!-- 表單標題 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:r>
                <w:t>員工資料表</w:t>
            </w:r>
        </w:p>

        <!-- 姓名欄位 -->
        <w:p>
            <w:r>
                <w:t>姓名：</w:t>
            </w:r>
            <w:sdt>
                <w:sdtPr>
                    <w:tag w:val="employee_name"/>
                    <w:id w:val="1"/>
                    <w:placeholder>
                        <w:docPart w:val="DefaultPlaceholder_Text"/>
                    </w:placeholder>
                    <w:text/>
                </w:sdtPr>
                <w:sdtContent>
                    <w:r>
                        <w:rPr>
                            <w:rStyle w:val="PlaceholderText"/>
                        </w:rPr>
                        <w:t>請輸入姓名</w:t>
                    </w:r>
                </w:sdtContent>
            </w:sdt>
        </w:p>

        <!-- 部門選擇 -->
        <w:p>
            <w:r>
                <w:t>部門：</w:t>
            </w:r>
            <w:sdt>
                <w:sdtPr>
                    <w:tag w:val="department"/>
                    <w:id w:val="2"/>
                    <w:dropDownList>
                        <w:listItem w:displayText="請選擇部門" w:value=""/>
                        <w:listItem w:displayText="研發部" w:value="RD"/>
                        <w:listItem w:displayText="業務部" w:value="SALES"/>
                        <w:listItem w:displayText="人資部" w:value="HR"/>
                    </w:dropDownList>
                </w:sdtPr>
                <w:sdtContent>
                    <w:r>
                        <w:t>請選擇部門</w:t>
                    </w:r>
                </w:sdtContent>
            </w:sdt>
        </w:p>

        <!-- 入職日期 -->
        <w:p>
            <w:r>
                <w:t>入職日期：</w:t>
            </w:r>
            <w:sdt>
                <w:sdtPr>
                    <w:tag w:val="hire_date"/>
                    <w:id w:val="3"/>
                    <w:date>
                        <w:dateFormat w:val="yyyy/MM/dd"/>
                        <w:lid w:val="zh-TW"/>
                    </w:date>
                </w:sdtPr>
                <w:sdtContent>
                    <w:r>
                        <w:t>請選擇日期</w:t>
                    </w:r>
                </w:sdtContent>
            </w:sdt>
        </w:p>

        <!-- 同意條款 -->
        <w:p>
            <w:sdt>
                <w:sdtPr>
                    <w:tag w:val="agree_terms"/>
                    <w:id w:val="4"/>
                    <w14:checkbox>
                        <w14:checked w14:val="0"/>
                        <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
                        <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
                    </w14:checkbox>
                </w:sdtPr>
                <w:sdtContent>
                    <w:r>
                        <w:rPr>
                            <w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/>
                        </w:rPr>
                        <w:t>☐</w:t>
                    </w:r>
                </w:sdtContent>
            </w:sdt>
            <w:r>
                <w:t> 我同意以上條款</w:t>
            </w:r>
        </w:p>

        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>
        </w:sectPr>
    </w:body>
</w:document>
```

---

## 表單保護

### 在 settings.xml 中設定

```xml
<w:settings>
    <w:documentProtection w:edit="forms" w:enforcement="1"/>
</w:settings>
```

### 保護模式

| edit 值 | 說明 |
|---------|------|
| `none` | 無保護 |
| `readOnly` | 唯讀 |
| `comments` | 只能加註解 |
| `trackedChanges` | 只能追蹤修訂 |
| `forms` | 只能填寫表單 |

---

## 實作注意事項

### SDT vs 傳統欄位
- SDT 是 Word 2007+ 的新格式
- 傳統欄位相容性更好但功能較少
- 建議新文件使用 SDT

### ID 唯一性
- 每個 SDT 的 `w:id` 必須唯一
- 使用隨機或遞增數字

### 佔位符文字
- 使用 `w:showingPlcHdr` 標記佔位符狀態
- 套用 PlaceholderText 樣式

### 命名空間
- 核取方塊需要 Word 2010 命名空間 (w14)
- 重複區段需要 Word 2012 命名空間 (w15)

---

## 相關連結

- [欄位代碼](64-fields.md)
- [樣式系統](30-styles.md)
- [段落結構](11-paragraph.md)
