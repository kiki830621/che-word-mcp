# 圖片與繪圖 (Images & Drawings)

## 概述

OOXML 使用 DrawingML 來處理圖片和圖形。圖片存儲在 `word/media/` 目錄中，透過關係參照。

## 檔案結構

```
word/
├── document.xml           # 主文件（包含 drawing 元素）
├── media/                 # 媒體檔案目錄
│   ├── image1.png
│   ├── image2.jpeg
│   └── image3.gif
└── _rels/
    └── document.xml.rels  # 關係（包含圖片參照）
```

---

## 關係設定

### document.xml.rels

```xml
<Relationship Id="rId5"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
<Relationship Id="rId6"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image2.jpeg"/>
```

### [Content_Types].xml

```xml
<Default Extension="png" ContentType="image/png"/>
<Default Extension="jpeg" ContentType="image/jpeg"/>
<Default Extension="jpg" ContentType="image/jpeg"/>
<Default Extension="gif" ContentType="image/gif"/>
<Default Extension="tiff" ContentType="image/tiff"/>
<Default Extension="bmp" ContentType="image/bmp"/>
<Default Extension="wmf" ContentType="image/x-wmf"/>
<Default Extension="emf" ContentType="image/x-emf"/>
```

---

## 圖片定位方式

| 類型 | 元素 | 說明 |
|------|------|------|
| 行內 | `wp:inline` | 隨文字流動 |
| 浮動 | `wp:anchor` | 定位在頁面上 |

---

## w:drawing（繪圖容器）

```xml
<w:r>
    <w:drawing>
        <!-- wp:inline 或 wp:anchor -->
    </w:drawing>
</w:r>
```

---

## wp:inline（行內圖片）

```xml
<w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0"
               xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">

        <!-- 尺寸 (EMU) -->
        <wp:extent cx="1905000" cy="1428750"/>

        <!-- 效果範圍 -->
        <wp:effectExtent l="0" t="0" r="0" b="0"/>

        <!-- 文件屬性 -->
        <wp:docPr id="1" name="Picture 1" descr="圖片描述"/>

        <!-- 非視覺圖形框架屬性 -->
        <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks noChangeAspect="1"
                                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        </wp:cNvGraphicFramePr>

        <!-- 圖形內容 -->
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <!-- ... 圖片詳細內容 ... -->
                </pic:pic>
            </a:graphicData>
        </a:graphic>
    </wp:inline>
</w:drawing>
```

### wp:inline 屬性

| 屬性 | 說明 |
|------|------|
| `distT` | 上方距離 (EMU) |
| `distB` | 下方距離 (EMU) |
| `distL` | 左側距離 (EMU) |
| `distR` | 右側距離 (EMU) |

---

## wp:anchor（浮動圖片）

```xml
<w:drawing>
    <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
               simplePos="0" relativeHeight="251658240"
               behindDoc="0" locked="0" layoutInCell="1"
               allowOverlap="1"
               xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">

        <!-- 簡單位置（通常不使用） -->
        <wp:simplePos x="0" y="0"/>

        <!-- 水平位置 -->
        <wp:positionH relativeFrom="column">
            <wp:posOffset>0</wp:posOffset>
            <!-- 或 -->
            <wp:align>center</wp:align>
        </wp:positionH>

        <!-- 垂直位置 -->
        <wp:positionV relativeFrom="paragraph">
            <wp:posOffset>0</wp:posOffset>
            <!-- 或 -->
            <wp:align>top</wp:align>
        </wp:positionV>

        <!-- 尺寸 -->
        <wp:extent cx="1905000" cy="1428750"/>

        <!-- 效果範圍 -->
        <wp:effectExtent l="0" t="0" r="0" b="0"/>

        <!-- 文繞圖設定 -->
        <wp:wrapSquare wrapText="bothSides"/>

        <!-- 文件屬性 -->
        <wp:docPr id="1" name="Picture 1"/>

        <!-- 非視覺框架屬性 -->
        <wp:cNvGraphicFramePr/>

        <!-- 圖形內容 -->
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <!-- ... -->
        </a:graphic>
    </wp:anchor>
</w:drawing>
```

### wp:anchor 屬性

| 屬性 | 說明 |
|------|------|
| `distT`, `distB`, `distL`, `distR` | 四周距離 (EMU) |
| `simplePos` | 使用簡單位置 |
| `relativeHeight` | 相對高度（Z 順序） |
| `behindDoc` | 在文字後方 |
| `locked` | 鎖定位置 |
| `layoutInCell` | 在表格儲存格內排版 |
| `allowOverlap` | 允許重疊 |

### 位置相對於 (relativeFrom)

| 值 | 說明 |
|----|------|
| `character` | 字元 |
| `column` | 欄 |
| `insideMargin` | 內側邊距 |
| `leftMargin` | 左邊距 |
| `line` | 行 |
| `margin` | 邊距 |
| `outsideMargin` | 外側邊距 |
| `page` | 頁面 |
| `paragraph` | 段落 |
| `rightMargin` | 右邊距 |
| `topMargin` | 上邊距 |
| `bottomMargin` | 下邊距 |

### 對齊 (wp:align)

| 值 | 說明 |
|----|------|
| `left` | 靠左 |
| `center` | 置中 |
| `right` | 靠右 |
| `top` | 靠上 |
| `bottom` | 靠下 |
| `inside` | 內側 |
| `outside` | 外側 |

---

## 文繞圖設定

### wp:wrapNone（無文繞圖）

```xml
<wp:wrapNone/>
```

### wp:wrapSquare（四方環繞）

```xml
<wp:wrapSquare wrapText="bothSides"/>
```

| wrapText | 說明 |
|----------|------|
| `bothSides` | 兩側環繞 |
| `left` | 僅左側 |
| `right` | 僅右側 |
| `largest` | 較大側 |

### wp:wrapTight（緊密環繞）

```xml
<wp:wrapTight wrapText="bothSides">
    <wp:wrapPolygon edited="0">
        <wp:start x="0" y="0"/>
        <wp:lineTo x="0" y="21600"/>
        <wp:lineTo x="21600" y="21600"/>
        <wp:lineTo x="21600" y="0"/>
        <wp:lineTo x="0" y="0"/>
    </wp:wrapPolygon>
</wp:wrapTight>
```

### wp:wrapThrough（穿越環繞）

```xml
<wp:wrapThrough wrapText="bothSides">
    <wp:wrapPolygon edited="0">
        <!-- 環繞多邊形 -->
    </wp:wrapPolygon>
</wp:wrapThrough>
```

### wp:wrapTopAndBottom（上下環繞）

```xml
<wp:wrapTopAndBottom/>
```

---

## pic:pic（圖片元素）

```xml
<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <!-- 非視覺圖片屬性 -->
    <pic:nvPicPr>
        <pic:cNvPr id="1" name="image1.png"/>
        <pic:cNvPicPr/>
    </pic:nvPicPr>

    <!-- 圖片填充 -->
    <pic:blipFill>
        <a:blip r:embed="rId5"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <!-- 可選：圖片效果 -->
        </a:blip>
        <a:stretch>
            <a:fillRect/>
        </a:stretch>
    </pic:blipFill>

    <!-- 形狀屬性 -->
    <pic:spPr>
        <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="1905000" cy="1428750"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
            <a:avLst/>
        </a:prstGeom>
    </pic:spPr>
</pic:pic>
```

---

## EMU 單位轉換

EMU (English Metric Units) 是 OOXML 中的基本長度單位。

```
1 inch = 914400 EMU
1 cm = 360000 EMU
1 pt = 12700 EMU
1 pixel (96 dpi) = 9525 EMU
```

### 常用尺寸換算

| 尺寸 | EMU |
|------|-----|
| 1 inch | 914400 |
| 2 inch | 1828800 |
| 5 cm | 1800000 |
| 10 cm | 3600000 |
| 100 px | 952500 |
| 200 px | 1905000 |

### 計算公式

```
EMU = pixels × 9525
EMU = inches × 914400
EMU = cm × 360000
EMU = pt × 12700
```

---

## 圖片效果

### 裁剪

```xml
<pic:blipFill>
    <a:blip r:embed="rId5"/>
    <a:srcRect l="10000" t="10000" r="10000" b="10000"/>
    <a:stretch>
        <a:fillRect/>
    </a:stretch>
</pic:blipFill>
```

`l`, `t`, `r`, `b` 是從各邊裁剪的百分比（千分比，10000 = 10%）。

### 旋轉

```xml
<pic:spPr>
    <a:xfrm rot="2700000">  <!-- 旋轉角度（60000 = 1度） -->
        <a:off x="0" y="0"/>
        <a:ext cx="1905000" cy="1428750"/>
    </a:xfrm>
</pic:spPr>
```

### 翻轉

```xml
<a:xfrm flipH="1">  <!-- 水平翻轉 -->
    <!-- ... -->
</a:xfrm>

<a:xfrm flipV="1">  <!-- 垂直翻轉 -->
    <!-- ... -->
</a:xfrm>
```

### 圖片效果濾鏡

```xml
<a:blip r:embed="rId5">
    <!-- 亮度/對比度 -->
    <a:lum bright="20000" contrast="10000"/>

    <!-- 灰階 -->
    <a:grayscl/>

    <!-- 雙色調 -->
    <a:duotone>
        <a:schemeClr val="accent1"/>
        <a:schemeClr val="accent2"/>
    </a:duotone>

    <!-- 色彩重新著色 -->
    <a:clrRepl>
        <a:srgbClr val="FF0000"/>
    </a:clrRepl>
</a:blip>
```

---

## 邊框和陰影

### 圖片邊框

```xml
<pic:spPr>
    <a:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="1905000" cy="1428750"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
        <a:avLst/>
    </a:prstGeom>
    <a:ln w="12700">
        <a:solidFill>
            <a:srgbClr val="000000"/>
        </a:solidFill>
    </a:ln>
</pic:spPr>
```

### 陰影效果

```xml
<pic:spPr>
    <!-- ... -->
    <a:effectLst>
        <a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl">
            <a:srgbClr val="000000">
                <a:alpha val="43000"/>
            </a:srgbClr>
        </a:outerShdw>
    </a:effectLst>
</pic:spPr>
```

---

## 完整範例

### 行內圖片

```xml
<w:p>
    <w:r>
        <w:t xml:space="preserve">這是一張圖片：</w:t>
    </w:r>
    <w:r>
        <w:drawing>
            <wp:inline distT="0" distB="0" distL="0" distR="0"
                       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <wp:extent cx="1905000" cy="1428750"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:docPr id="1" name="Picture 1" descr="範例圖片"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic>
                            <pic:nvPicPr>
                                <pic:cNvPr id="1" name="image1.png"/>
                                <pic:cNvPicPr/>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="rId5"/>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="1905000" cy="1428750"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:inline>
        </w:drawing>
    </w:r>
</w:p>
```

### 浮動圖片（置中）

```xml
<w:p>
    <w:r>
        <w:drawing>
            <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
                       simplePos="0" relativeHeight="251658240"
                       behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"
                       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <wp:simplePos x="0" y="0"/>
                <wp:positionH relativeFrom="margin">
                    <wp:align>center</wp:align>
                </wp:positionH>
                <wp:positionV relativeFrom="paragraph">
                    <wp:posOffset>0</wp:posOffset>
                </wp:positionV>
                <wp:extent cx="3810000" cy="2857500"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:wrapTopAndBottom/>
                <wp:docPr id="2" name="Picture 2"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic>
                            <pic:nvPicPr>
                                <pic:cNvPr id="2" name="image2.png"/>
                                <pic:cNvPicPr/>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="rId6"/>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="3810000" cy="2857500"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:anchor>
        </w:drawing>
    </w:r>
</w:p>
```

---

## 下一步

- [51-hyperlinks.md](51-hyperlinks.md) - 超連結
- [60-comments.md](60-comments.md) - 註解
- [64-fields.md](64-fields.md) - 欄位代碼
