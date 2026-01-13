# 目錄 (Table of Contents)

## 概述

目錄（Table of Contents, TOC）是自動根據文件中的標題樣式生成的導覽結構。OOXML 使用欄位代碼（Field Code）來定義目錄，並包含快取的靜態內容以便顯示。

## 基本結構

目錄由以下元素組成：
1. **結構化文件標籤（SDT）**：包裝整個目錄
2. **欄位代碼**：定義目錄的生成規則
3. **快取內容**：已生成的目錄項目

---

## 簡單目錄

### 基本 TOC 欄位

```xml
<w:p>
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <!-- 這裡是快取的目錄內容 -->
    <w:r>
        <w:t>目錄項目會顯示在這裡</w:t>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
</w:p>
```

### TOC 開關參數

| 開關 | 說明 | 範例 |
|------|------|------|
| `\o "1-3"` | 大綱層級範圍 | 包含標題 1-3 |
| `\h` | 超連結 | 目錄項目可點擊 |
| `\z` | 隱藏頁碼（Web 檢視） | |
| `\u` | 使用段落大綱層級 | |
| `\t "樣式,層級"` | 指定樣式對應 | `\t "MyHeading,1"` |
| `\n "範圍"` | 省略頁碼 | `\n "1-1"` 省略標題1頁碼 |
| `\p "分隔符"` | 頁碼分隔符 | `\p "-"` |
| `\w` | 保留 Tab 字元 | |
| `\x` | 保留換行符 | |
| `\c "序列"` | 使用 SEQ 欄位序列 | 圖表目錄用 |

---

## 使用 SDT 包裝的目錄

### 完整結構

```xml
<w:sdt>
    <w:sdtPr>
        <w:docPartObj>
            <w:docPartGallery w:val="Table of Contents"/>
            <w:docPartUnique/>
        </w:docPartObj>
    </w:sdtPr>
    <w:sdtEndPr/>
    <w:sdtContent>
        <!-- 目錄標題 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="TOCHeading"/>
            </w:pPr>
            <w:r>
                <w:t>目錄</w:t>
            </w:r>
        </w:p>

        <!-- TOC 欄位開始 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="TOC1"/>
                <w:tabs>
                    <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
                </w:tabs>
            </w:pPr>
            <w:r>
                <w:fldChar w:fldCharType="begin"/>
            </w:r>
            <w:r>
                <w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>
            </w:r>
            <w:r>
                <w:fldChar w:fldCharType="separate"/>
            </w:r>
        </w:p>

        <!-- 快取的目錄項目 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="TOC1"/>
            </w:pPr>
            <w:hyperlink w:anchor="_Toc123456781">
                <w:r>
                    <w:t>第一章 緒論</w:t>
                </w:r>
                <w:r>
                    <w:tab/>
                </w:r>
                <w:r>
                    <w:t>1</w:t>
                </w:r>
            </w:hyperlink>
        </w:p>

        <w:p>
            <w:pPr>
                <w:pStyle w:val="TOC2"/>
            </w:pPr>
            <w:hyperlink w:anchor="_Toc123456782">
                <w:r>
                    <w:t>1.1 研究背景</w:t>
                </w:r>
                <w:r>
                    <w:tab/>
                </w:r>
                <w:r>
                    <w:t>2</w:t>
                </w:r>
            </w:hyperlink>
        </w:p>

        <!-- TOC 欄位結束 -->
        <w:p>
            <w:r>
                <w:fldChar w:fldCharType="end"/>
            </w:r>
        </w:p>

    </w:sdtContent>
</w:sdt>
```

---

## 目錄樣式

### styles.xml 中的 TOC 樣式

```xml
<!-- 目錄標題 -->
<w:style w:type="paragraph" w:styleId="TOCHeading">
    <w:name w:val="TOC Heading"/>
    <w:basedOn w:val="Heading1"/>
    <w:next w:val="Normal"/>
    <w:pPr>
        <w:outlineLvl w:val="9"/>  <!-- 不出現在目錄中 -->
    </w:pPr>
</w:style>

<!-- TOC 層級 1 -->
<w:style w:type="paragraph" w:styleId="TOC1">
    <w:name w:val="toc 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:autoRedefine/>
    <w:pPr>
        <w:tabs>
            <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
        </w:tabs>
        <w:spacing w:after="100"/>
    </w:pPr>
</w:style>

<!-- TOC 層級 2 -->
<w:style w:type="paragraph" w:styleId="TOC2">
    <w:name w:val="toc 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:autoRedefine/>
    <w:pPr>
        <w:tabs>
            <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
        </w:tabs>
        <w:spacing w:after="100"/>
        <w:ind w:left="220"/>
    </w:pPr>
</w:style>

<!-- TOC 層級 3 -->
<w:style w:type="paragraph" w:styleId="TOC3">
    <w:name w:val="toc 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:autoRedefine/>
    <w:pPr>
        <w:tabs>
            <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
        </w:tabs>
        <w:spacing w:after="100"/>
        <w:ind w:left="440"/>
    </w:pPr>
</w:style>

<!-- 超連結樣式 -->
<w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr>
        <w:color w:val="0563C1" w:themeColor="hyperlink"/>
        <w:u w:val="single"/>
    </w:rPr>
</w:style>
```

---

## 標題書籤

### 為標題加入書籤

目錄的超連結需要對應標題的書籤：

```xml
<!-- 標題 1 加入書籤 -->
<w:p>
    <w:pPr>
        <w:pStyle w:val="Heading1"/>
    </w:pPr>
    <w:bookmarkStart w:id="0" w:name="_Toc123456781"/>
    <w:r>
        <w:t>第一章 緒論</w:t>
    </w:r>
    <w:bookmarkEnd w:id="0"/>
</w:p>

<!-- 標題 2 加入書籤 -->
<w:p>
    <w:pPr>
        <w:pStyle w:val="Heading2"/>
    </w:pPr>
    <w:bookmarkStart w:id="1" w:name="_Toc123456782"/>
    <w:r>
        <w:t>1.1 研究背景</w:t>
    </w:r>
    <w:bookmarkEnd w:id="1"/>
</w:p>
```

### 書籤命名規則
- TOC 書籤通常以 `_Toc` 開頭
- 後接唯一數字識別碼
- 範例：`_Toc123456781`, `_Toc123456782`

---

## 圖表目錄

### 圖表目錄欄位

```xml
<w:p>
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText xml:space="preserve"> TOC \h \z \c "Figure" </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <!-- 快取的圖表目錄 -->
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
</w:p>
```

### 圖表標題使用 SEQ 欄位

```xml
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

---

## 表格目錄

```xml
<w:p>
    <w:r>
        <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
        <w:instrText xml:space="preserve"> TOC \h \z \c "Table" </w:instrText>
    </w:r>
    <w:r>
        <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <!-- 快取的表格目錄 -->
    <w:r>
        <w:fldChar w:fldCharType="end"/>
    </w:r>
</w:p>
```

---

## 完整範例

### document.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>

        <!-- 目錄區塊 -->
        <w:sdt>
            <w:sdtPr>
                <w:docPartObj>
                    <w:docPartGallery w:val="Table of Contents"/>
                    <w:docPartUnique/>
                </w:docPartObj>
            </w:sdtPr>
            <w:sdtContent>
                <!-- 目錄標題 -->
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="TOCHeading"/>
                    </w:pPr>
                    <w:r>
                        <w:t>目錄</w:t>
                    </w:r>
                </w:p>

                <!-- TOC 欄位 -->
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="TOC1"/>
                    </w:pPr>
                    <w:r>
                        <w:fldChar w:fldCharType="begin"/>
                    </w:r>
                    <w:r>
                        <w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>
                    </w:r>
                    <w:r>
                        <w:fldChar w:fldCharType="separate"/>
                    </w:r>
                </w:p>

                <!-- 目錄項目 1 -->
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="TOC1"/>
                        <w:tabs>
                            <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
                        </w:tabs>
                    </w:pPr>
                    <w:hyperlink w:anchor="_Toc001">
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:t>第一章 緒論</w:t>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:tab/>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:fldChar w:fldCharType="begin"/>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:instrText xml:space="preserve"> PAGEREF _Toc001 \h </w:instrText>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:fldChar w:fldCharType="separate"/>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:t>1</w:t>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:fldChar w:fldCharType="end"/>
                        </w:r>
                    </w:hyperlink>
                </w:p>

                <!-- 目錄項目 2 (層級 2) -->
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="TOC2"/>
                        <w:tabs>
                            <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
                        </w:tabs>
                    </w:pPr>
                    <w:hyperlink w:anchor="_Toc002">
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:t>1.1 研究背景</w:t>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:tab/>
                        </w:r>
                        <w:r>
                            <w:rPr>
                                <w:noProof/>
                            </w:rPr>
                            <w:t>2</w:t>
                        </w:r>
                    </w:hyperlink>
                </w:p>

                <!-- TOC 欄位結束 -->
                <w:p>
                    <w:r>
                        <w:fldChar w:fldCharType="end"/>
                    </w:r>
                </w:p>

            </w:sdtContent>
        </w:sdt>

        <!-- 分頁 -->
        <w:p>
            <w:r>
                <w:br w:type="page"/>
            </w:r>
        </w:p>

        <!-- 正文：標題 1 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:bookmarkStart w:id="0" w:name="_Toc001"/>
            <w:r>
                <w:t>第一章 緒論</w:t>
            </w:r>
            <w:bookmarkEnd w:id="0"/>
        </w:p>

        <!-- 正文：標題 2 -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading2"/>
            </w:pPr>
            <w:bookmarkStart w:id="1" w:name="_Toc002"/>
            <w:r>
                <w:t>1.1 研究背景</w:t>
            </w:r>
            <w:bookmarkEnd w:id="1"/>
        </w:p>

        <w:p>
            <w:r>
                <w:t>本研究探討...</w:t>
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

## 更新目錄

### 更新機制
- OOXML 中的目錄是靜態快取
- Word 開啟時會提示更新
- 也可使用 VBA 或其他工具程式更新

### 更新標記

```xml
<w:settings>
    <!-- 開啟時更新欄位 -->
    <w:updateFields w:val="true"/>
</w:settings>
```

---

## 實作注意事項

### 欄位結構
1. `fldChar begin` - 欄位開始
2. `instrText` - 欄位指令
3. `fldChar separate` - 分隔（指令與結果）
4. 快取內容 - 顯示的結果
5. `fldChar end` - 欄位結束

### 書籤同步
- 目錄超連結的 `w:anchor` 必須對應正確的書籤名稱
- 新增/刪除標題時需更新書籤

### 樣式一致性
- 目錄項目使用 TOC1, TOC2, TOC3... 樣式
- 標題使用 Heading1, Heading2, Heading3... 樣式
- 確保大綱層級（outlineLvl）正確設定

---

## 相關連結

- [欄位代碼](64-fields.md)
- [書籤](51-bookmarks.md)
- [樣式系統](30-styles.md)
- [段落格式](14-paragraph-formatting.md)
