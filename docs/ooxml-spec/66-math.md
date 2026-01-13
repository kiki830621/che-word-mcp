# 數學公式 (Math / OMML)

## 概述

OOXML 使用 Office Math Markup Language (OMML) 來表示數學公式。OMML 是一種 XML 格式，可以表達複雜的數學表達式，包括分數、根號、積分、矩陣等。

## 命名空間

```xml
xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
```

## 基本結構

### 數學區塊

```xml
<w:p>
    <m:oMath>
        <!-- 數學內容 -->
    </m:oMath>
</w:p>
```

### 獨立數學段落

```xml
<w:p>
    <m:oMathPara>
        <m:oMathParaPr>
            <m:jc m:val="center"/>  <!-- 對齊方式 -->
        </m:oMathParaPr>
        <m:oMath>
            <!-- 數學內容 -->
        </m:oMath>
    </m:oMathPara>
</w:p>
```

---

## 基本元素

### 文字 (m:r)

```xml
<m:oMath>
    <m:r>
        <m:t>x</m:t>
    </m:r>
    <m:r>
        <m:t>+</m:t>
    </m:r>
    <m:r>
        <m:t>y</m:t>
    </m:r>
</m:oMath>
<!-- 結果：x+y -->
```

### 數字與運算子

```xml
<m:oMath>
    <m:r>
        <m:t>2</m:t>
    </m:r>
    <m:r>
        <m:t>×</m:t>
    </m:r>
    <m:r>
        <m:t>3</m:t>
    </m:r>
    <m:r>
        <m:t>=</m:t>
    </m:r>
    <m:r>
        <m:t>6</m:t>
    </m:r>
</m:oMath>
<!-- 結果：2×3=6 -->
```

---

## 分數 (m:f)

### 基本分數

```xml
<m:f>
    <m:fPr>
        <m:type m:val="bar"/>  <!-- 分數線類型 -->
    </m:fPr>
    <m:num>  <!-- 分子 -->
        <m:r>
            <m:t>1</m:t>
        </m:r>
    </m:num>
    <m:den>  <!-- 分母 -->
        <m:r>
            <m:t>2</m:t>
        </m:r>
    </m:den>
</m:f>
<!-- 結果：½ -->
```

### 分數類型 (m:type)

| 值 | 說明 | 範例 |
|----|------|------|
| `bar` | 標準分數線 | ½ |
| `skw` | 斜線分數 | 1/2 |
| `lin` | 線性分數 | 1/2 |
| `noBar` | 無分數線（堆疊） | ¹₂ |

### 複雜分數

```xml
<m:f>
    <m:num>
        <m:r><m:t>a</m:t></m:r>
        <m:r><m:t>+</m:t></m:r>
        <m:r><m:t>b</m:t></m:r>
    </m:num>
    <m:den>
        <m:r><m:t>c</m:t></m:r>
        <m:r><m:t>-</m:t></m:r>
        <m:r><m:t>d</m:t></m:r>
    </m:den>
</m:f>
<!-- 結果：(a+b)/(c-d) -->
```

---

## 上標與下標

### 上標 (m:sSup)

```xml
<m:sSup>
    <m:e>  <!-- 基底 -->
        <m:r><m:t>x</m:t></m:r>
    </m:e>
    <m:sup>  <!-- 上標 -->
        <m:r><m:t>2</m:t></m:r>
    </m:sup>
</m:sSup>
<!-- 結果：x² -->
```

### 下標 (m:sSub)

```xml
<m:sSub>
    <m:e>  <!-- 基底 -->
        <m:r><m:t>a</m:t></m:r>
    </m:e>
    <m:sub>  <!-- 下標 -->
        <m:r><m:t>n</m:t></m:r>
    </m:sub>
</m:sSub>
<!-- 結果：aₙ -->
```

### 上下標 (m:sSubSup)

```xml
<m:sSubSup>
    <m:e>
        <m:r><m:t>x</m:t></m:r>
    </m:e>
    <m:sub>
        <m:r><m:t>i</m:t></m:r>
    </m:sub>
    <m:sup>
        <m:r><m:t>2</m:t></m:r>
    </m:sup>
</m:sSubSup>
<!-- 結果：xᵢ² -->
```

---

## 根號 (m:rad)

### 平方根

```xml
<m:rad>
    <m:radPr>
        <m:degHide m:val="1"/>  <!-- 隱藏根次 -->
    </m:radPr>
    <m:deg/>  <!-- 空的根次 -->
    <m:e>
        <m:r><m:t>x</m:t></m:r>
    </m:e>
</m:rad>
<!-- 結果：√x -->
```

### n 次方根

```xml
<m:rad>
    <m:deg>
        <m:r><m:t>3</m:t></m:r>
    </m:deg>
    <m:e>
        <m:r><m:t>x</m:t></m:r>
    </m:e>
</m:rad>
<!-- 結果：∛x -->
```

### 複雜根號

```xml
<m:rad>
    <m:radPr>
        <m:degHide m:val="1"/>
    </m:radPr>
    <m:deg/>
    <m:e>
        <m:sSup>
            <m:e><m:r><m:t>a</m:t></m:r></m:e>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup>
        <m:r><m:t>+</m:t></m:r>
        <m:sSup>
            <m:e><m:r><m:t>b</m:t></m:r></m:e>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup>
    </m:e>
</m:rad>
<!-- 結果：√(a²+b²) -->
```

---

## 積分與極限

### 積分 (m:nary)

```xml
<m:nary>
    <m:naryPr>
        <m:chr m:val="∫"/>  <!-- 積分符號 -->
        <m:limLoc m:val="subSup"/>  <!-- 上下限位置 -->
    </m:naryPr>
    <m:sub>  <!-- 下限 -->
        <m:r><m:t>0</m:t></m:r>
    </m:sub>
    <m:sup>  <!-- 上限 -->
        <m:r><m:t>1</m:t></m:r>
    </m:sup>
    <m:e>  <!-- 被積函數 -->
        <m:r><m:t>f(x)dx</m:t></m:r>
    </m:e>
</m:nary>
<!-- 結果：∫₀¹ f(x)dx -->
```

### 常用 N-ary 符號

| 符號 | Unicode | 說明 |
|------|---------|------|
| ∫ | 222B | 積分 |
| ∬ | 222C | 雙重積分 |
| ∭ | 222D | 三重積分 |
| ∮ | 222E | 環路積分 |
| ∑ | 2211 | 求和 |
| ∏ | 220F | 乘積 |
| ⋃ | 22C3 | 聯集 |
| ⋂ | 22C2 | 交集 |

### 求和

```xml
<m:nary>
    <m:naryPr>
        <m:chr m:val="∑"/>
        <m:limLoc m:val="undOvr"/>  <!-- 上下標在正上下方 -->
    </m:naryPr>
    <m:sub>
        <m:r><m:t>i=1</m:t></m:r>
    </m:sub>
    <m:sup>
        <m:r><m:t>n</m:t></m:r>
    </m:sup>
    <m:e>
        <m:sSub>
            <m:e><m:r><m:t>a</m:t></m:r></m:e>
            <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
        </m:sSub>
    </m:e>
</m:nary>
<!-- 結果：∑ᵢ₌₁ⁿ aᵢ -->
```

### 極限

```xml
<m:func>
    <m:funcPr>
        <!-- 函數屬性 -->
    </m:funcPr>
    <m:fName>
        <m:limLow>
            <m:e>
                <m:r>
                    <m:rPr>
                        <m:scr m:val="roman"/>
                    </m:rPr>
                    <m:t>lim</m:t>
                </m:r>
            </m:e>
            <m:lim>
                <m:r><m:t>x→∞</m:t></m:r>
            </m:lim>
        </m:limLow>
    </m:fName>
    <m:e>
        <m:r><m:t>f(x)</m:t></m:r>
    </m:e>
</m:func>
<!-- 結果：lim(x→∞) f(x) -->
```

---

## 括號 (m:d)

### 基本括號

```xml
<m:d>
    <m:dPr>
        <m:begChr m:val="("/>
        <m:endChr m:val=")"/>
    </m:dPr>
    <m:e>
        <m:r><m:t>a+b</m:t></m:r>
    </m:e>
</m:d>
<!-- 結果：(a+b) -->
```

### 括號類型

| 開始 | 結束 | 說明 |
|------|------|------|
| ( | ) | 小括號 |
| [ | ] | 方括號 |
| { | } | 大括號 |
| | | | | 絕對值 |
| ⌈ | ⌉ | 上取整 |
| ⌊ | ⌋ | 下取整 |
| 〈 | 〉 | 角括號 |

### 絕對值

```xml
<m:d>
    <m:dPr>
        <m:begChr m:val="|"/>
        <m:endChr m:val="|"/>
    </m:dPr>
    <m:e>
        <m:r><m:t>x</m:t></m:r>
    </m:e>
</m:d>
<!-- 結果：|x| -->
```

---

## 矩陣 (m:m)

### 基本矩陣

```xml
<m:d>
    <m:dPr>
        <m:begChr m:val="["/>
        <m:endChr m:val="]"/>
    </m:dPr>
    <m:e>
        <m:m>
            <m:mPr>
                <m:mcs>
                    <m:mc>
                        <m:mcPr>
                            <m:count m:val="2"/>
                            <m:mcJc m:val="center"/>
                        </m:mcPr>
                    </m:mc>
                </m:mcs>
            </m:mPr>
            <!-- 第一列 -->
            <m:mr>
                <m:e><m:r><m:t>a</m:t></m:r></m:e>
                <m:e><m:r><m:t>b</m:t></m:r></m:e>
            </m:mr>
            <!-- 第二列 -->
            <m:mr>
                <m:e><m:r><m:t>c</m:t></m:r></m:e>
                <m:e><m:r><m:t>d</m:t></m:r></m:e>
            </m:mr>
        </m:m>
    </m:e>
</m:d>
<!-- 結果：[a b; c d] -->
```

### 矩陣對齊

| m:mcJc 值 | 說明 |
|-----------|------|
| `left` | 靠左 |
| `center` | 置中 |
| `right` | 靠右 |

---

## 函數名稱

### 三角函數

```xml
<m:func>
    <m:fName>
        <m:r>
            <m:rPr>
                <m:scr m:val="roman"/>  <!-- 正體 -->
            </m:rPr>
            <m:t>sin</m:t>
        </m:r>
    </m:fName>
    <m:e>
        <m:r><m:t>θ</m:t></m:r>
    </m:e>
</m:func>
<!-- 結果：sin θ -->
```

### 常用函數

| 函數 | 說明 |
|------|------|
| sin, cos, tan | 三角函數 |
| sec, csc, cot | 餘割等 |
| arcsin, arccos | 反三角 |
| sinh, cosh, tanh | 雙曲函數 |
| log, ln, lg | 對數 |
| exp | 指數 |
| max, min | 最大最小 |
| gcd, lcm | 最大公因數 |
| det | 行列式 |
| dim | 維度 |

---

## 重音與修飾

### 上方符號 (m:acc)

```xml
<m:acc>
    <m:accPr>
        <m:chr m:val="̂"/>  <!-- 帽子 -->
    </m:accPr>
    <m:e>
        <m:r><m:t>x</m:t></m:r>
    </m:e>
</m:acc>
<!-- 結果：x̂ -->
```

### 常用重音符號

| 符號 | Unicode | 說明 |
|------|---------|------|
| ̂ | 0302 | 帽子 |
| ̄ | 0304 | 橫線 |
| ̃ | 0303 | 波浪 |
| ⃗ | 20D7 | 向量箭頭 |
| ̇ | 0307 | 點 |
| ̈ | 0308 | 雙點 |

### 上橫線

```xml
<m:bar>
    <m:barPr>
        <m:pos m:val="top"/>
    </m:barPr>
    <m:e>
        <m:r><m:t>x</m:t></m:r>
    </m:e>
</m:bar>
<!-- 結果：x̄ -->
```

---

## 完整範例：二次方程式

```xml
<w:p>
    <m:oMathPara>
        <m:oMathParaPr>
            <m:jc m:val="center"/>
        </m:oMathParaPr>
        <m:oMath>
            <!-- x = -->
            <m:r><m:t>x=</m:t></m:r>

            <!-- 分數 -->
            <m:f>
                <m:num>
                    <!-- -b ± √(b²-4ac) -->
                    <m:r><m:t>-b±</m:t></m:r>
                    <m:rad>
                        <m:radPr>
                            <m:degHide m:val="1"/>
                        </m:radPr>
                        <m:deg/>
                        <m:e>
                            <m:sSup>
                                <m:e><m:r><m:t>b</m:t></m:r></m:e>
                                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                            </m:sSup>
                            <m:r><m:t>-4ac</m:t></m:r>
                        </m:e>
                    </m:rad>
                </m:num>
                <m:den>
                    <!-- 2a -->
                    <m:r><m:t>2a</m:t></m:r>
                </m:den>
            </m:f>
        </m:oMath>
    </m:oMathPara>
</w:p>
```

---

## 數學設定

### 在 settings.xml 中

```xml
<m:mathPr xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <m:mathFont m:val="Cambria Math"/>
    <m:brkBin m:val="before"/>
    <m:brkBinSub m:val="--"/>
    <m:smallFrac m:val="0"/>
    <m:dispDef/>
    <m:lMargin m:val="0"/>
    <m:rMargin m:val="0"/>
    <m:defJc m:val="centerGroup"/>
    <m:wrapIndent m:val="1440"/>
    <m:intLim m:val="subSup"/>
    <m:naryLim m:val="undOvr"/>
</m:mathPr>
```

### 設定項目說明

| 元素 | 說明 |
|------|------|
| `m:mathFont` | 數學字型 |
| `m:brkBin` | 二元運算子斷行（before/after） |
| `m:smallFrac` | 使用小型分數 |
| `m:dispDef` | 顯示模式預設 |
| `m:defJc` | 預設對齊 |
| `m:intLim` | 積分極限位置 |
| `m:naryLim` | N-ary 極限位置 |

---

## 文字屬性 (m:rPr)

```xml
<m:r>
    <m:rPr>
        <m:scr m:val="roman"/>  <!-- 字體樣式 -->
        <m:sty m:val="bi"/>     <!-- 粗斜體 -->
    </m:rPr>
    <m:t>sin</m:t>
</m:r>
```

### 字體樣式 (m:scr)

| 值 | 說明 |
|----|------|
| `roman` | 正體 |
| `script` | 手寫體 |
| `fraktur` | 哥德體 |
| `double-struck` | 雙線體（黑板粗體） |
| `sans-serif` | 無襯線 |
| `monospace` | 等寬 |

### 文字樣式 (m:sty)

| 值 | 說明 |
|----|------|
| `p` | 正體 (plain) |
| `b` | 粗體 |
| `i` | 斜體 |
| `bi` | 粗斜體 |

---

## 實作注意事項

### 字型
- 預設使用 Cambria Math 字型
- 確保系統有安裝數學字型

### 相容性
- OMML 是 Word 2007+ 的原生格式
- 也可轉換為 MathML 或 LaTeX

### 渲染
- 數學公式需要特殊渲染引擎
- 直接顯示 XML 無法正確呈現

---

## 相關連結

- [欄位代碼](64-fields.md)
- [段落結構](11-paragraph.md)
- [文字格式](13-text-formatting.md)
