import Foundation

// MARK: - Table of Contents (TOC)

/// 目錄欄位
struct TableOfContents {
    var title: String?                  // 目錄標題
    var headingLevels: ClosedRange<Int> // 包含的標題層級 (1-9)
    var includePageNumbers: Bool        // 是否包含頁碼
    var rightAlignPageNumbers: Bool     // 頁碼是否右對齊
    var useHyperlinks: Bool             // 是否使用超連結
    var tabLeader: TabLeader            // 定位線類型

    init(
        title: String? = nil,
        headingLevels: ClosedRange<Int> = 1...3,
        includePageNumbers: Bool = true,
        rightAlignPageNumbers: Bool = true,
        useHyperlinks: Bool = true,
        tabLeader: TabLeader = .dot
    ) {
        self.title = title
        self.headingLevels = headingLevels
        self.includePageNumbers = includePageNumbers
        self.rightAlignPageNumbers = rightAlignPageNumbers
        self.useHyperlinks = useHyperlinks
        self.tabLeader = tabLeader
    }
}

/// 定位線類型
enum TabLeader: String, Codable {
    case none = "none"
    case dot = "dot"
    case hyphen = "hyphen"
    case underscore = "underscore"
}

// MARK: - TOC XML Generation

extension TableOfContents {
    /// 產生目錄欄位的 XML
    func toXML() -> String {
        var xml = ""

        // 目錄標題（如果有）
        if let title = title {
            xml += """
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TOCHeading"/>
                </w:pPr>
                <w:r>
                    <w:t>\(escapeXML(title))</w:t>
                </w:r>
            </w:p>
            """
        }

        // 目錄開始 - 使用 SDT (Structured Document Tag)
        xml += """
        <w:sdt>
            <w:sdtPr>
                <w:docPartObj>
                    <w:docPartGallery w:val="Table of Contents"/>
                    <w:docPartUnique/>
                </w:docPartObj>
            </w:sdtPr>
            <w:sdtContent>
        """

        // 目錄欄位開始
        xml += """
            <w:p>
                <w:r>
                    <w:fldChar w:fldCharType="begin"/>
                </w:r>
                <w:r>
                    <w:instrText xml:space="preserve"> TOC \\o "\(headingLevels.lowerBound)-\(headingLevels.upperBound)"</w:instrText>
                </w:r>
        """

        // 欄位選項
        if useHyperlinks {
            xml += """
                <w:r>
                    <w:instrText xml:space="preserve"> \\h</w:instrText>
                </w:r>
            """
        }

        if !includePageNumbers {
            xml += """
                <w:r>
                    <w:instrText xml:space="preserve"> \\n</w:instrText>
                </w:r>
            """
        }

        // 欄位分隔和結束
        xml += """
                <w:r>
                    <w:fldChar w:fldCharType="separate"/>
                </w:r>
                <w:r>
                    <w:t>Update this field to generate table of contents.</w:t>
                </w:r>
                <w:r>
                    <w:fldChar w:fldCharType="end"/>
                </w:r>
            </w:p>
        """

        // SDT 結束
        xml += """
            </w:sdtContent>
        </w:sdt>
        """

        return xml
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - Form Controls

/// 表單文字欄位
struct FormTextField {
    var name: String                // 欄位名稱
    var defaultValue: String?       // 預設值
    var maxLength: Int?             // 最大長度
    var helpText: String?           // 說明文字

    init(name: String, defaultValue: String? = nil, maxLength: Int? = nil, helpText: String? = nil) {
        self.name = name
        self.defaultValue = defaultValue
        self.maxLength = maxLength
        self.helpText = helpText
    }
}

extension FormTextField {
    /// 產生表單文字欄位的 XML
    func toXML() -> String {
        var xml = "<w:sdt>"

        // SDT 屬性
        xml += "<w:sdtPr>"
        xml += "<w:alias w:val=\"\(escapeXML(name))\"/>"
        xml += "<w:tag w:val=\"\(escapeXML(name))\"/>"
        xml += "<w:text/>"
        xml += "</w:sdtPr>"

        // SDT 內容
        xml += "<w:sdtContent>"
        xml += "<w:r>"
        if let value = defaultValue {
            xml += "<w:t>\(escapeXML(value))</w:t>"
        } else {
            xml += "<w:t>          </w:t>"  // 空白佔位
        }
        xml += "</w:r>"
        xml += "</w:sdtContent>"

        xml += "</w:sdt>"
        return xml
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

/// 核取方塊
struct FormCheckbox {
    var name: String            // 欄位名稱
    var isChecked: Bool         // 是否勾選
    var checkedSymbol: String   // 勾選時的符號
    var uncheckedSymbol: String // 未勾選時的符號

    init(name: String, isChecked: Bool = false, checkedSymbol: String = "☒", uncheckedSymbol: String = "☐") {
        self.name = name
        self.isChecked = isChecked
        self.checkedSymbol = checkedSymbol
        self.uncheckedSymbol = uncheckedSymbol
    }
}

extension FormCheckbox {
    /// 產生核取方塊的 XML
    func toXML() -> String {
        var xml = "<w:sdt>"

        // SDT 屬性
        xml += "<w:sdtPr>"
        xml += "<w:alias w:val=\"\(escapeXML(name))\"/>"
        xml += "<w:tag w:val=\"\(escapeXML(name))\"/>"
        xml += "<w14:checkbox xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">"
        xml += "<w14:checked w14:val=\"\(isChecked ? "1" : "0")\"/>"
        xml += "<w14:checkedState w14:val=\"2612\" w14:font=\"MS Gothic\"/>"
        xml += "<w14:uncheckedState w14:val=\"2610\" w14:font=\"MS Gothic\"/>"
        xml += "</w14:checkbox>"
        xml += "</w:sdtPr>"

        // SDT 內容
        xml += "<w:sdtContent>"
        xml += "<w:r>"
        xml += "<w:rPr><w:rFonts w:ascii=\"MS Gothic\" w:hAnsi=\"MS Gothic\" w:hint=\"eastAsia\"/></w:rPr>"
        xml += "<w:t>\(isChecked ? "☒" : "☐")</w:t>"
        xml += "</w:r>"
        xml += "</w:sdtContent>"

        xml += "</w:sdt>"
        return xml
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

/// 下拉選單
struct FormDropdown {
    var name: String            // 欄位名稱
    var options: [String]       // 選項列表
    var selectedIndex: Int      // 選中的索引

    init(name: String, options: [String], selectedIndex: Int = 0) {
        self.name = name
        self.options = options
        self.selectedIndex = min(selectedIndex, options.count - 1)
    }
}

extension FormDropdown {
    /// 產生下拉選單的 XML
    func toXML() -> String {
        var xml = "<w:sdt>"

        // SDT 屬性
        xml += "<w:sdtPr>"
        xml += "<w:alias w:val=\"\(escapeXML(name))\"/>"
        xml += "<w:tag w:val=\"\(escapeXML(name))\"/>"
        xml += "<w:dropDownList>"
        for (index, option) in options.enumerated() {
            xml += "<w:listItem w:displayText=\"\(escapeXML(option))\" w:value=\"\(index)\"/>"
        }
        xml += "</w:dropDownList>"
        xml += "</w:sdtPr>"

        // SDT 內容
        xml += "<w:sdtContent>"
        xml += "<w:r>"
        let selectedValue = options.indices.contains(selectedIndex) ? options[selectedIndex] : ""
        xml += "<w:t>\(escapeXML(selectedValue))</w:t>"
        xml += "</w:r>"
        xml += "</w:sdtContent>"

        xml += "</w:sdt>"
        return xml
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - Mathematical Equations (OMML)

/// Office 數學公式
struct MathEquation {
    var latex: String           // LaTeX 格式的公式
    var displayMode: Bool       // 是否為獨立區塊（true）或行內（false）

    init(latex: String, displayMode: Bool = false) {
        self.latex = latex
        self.displayMode = displayMode
    }
}

extension MathEquation {
    /// 產生 OMML 格式的數學公式 XML
    /// 注意：這是簡化版本，僅支援基本公式
    func toXML() -> String {
        // OMML (Office Math Markup Language) 基本結構
        var xml = ""

        if displayMode {
            // 獨立區塊公式
            xml += "<w:p>"
            xml += "<w:pPr><w:jc w:val=\"center\"/></w:pPr>"
        }

        xml += "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"

        // 簡單轉換 LaTeX 到 OMML
        // 這是一個基礎實作，僅支援簡單文字
        let processedText = processLatex(latex)
        xml += "<m:r><m:t>\(escapeXML(processedText))</m:t></m:r>"

        xml += "</m:oMath>"

        if displayMode {
            xml += "</w:p>"
        }

        return xml
    }

    /// 基礎 LaTeX 處理（簡化版）
    private func processLatex(_ latex: String) -> String {
        var result = latex
        // 移除一些 LaTeX 指令，保留內容
        result = result.replacingOccurrences(of: "\\frac{", with: "(")
        result = result.replacingOccurrences(of: "}{", with: ")/(")
        result = result.replacingOccurrences(of: "\\sqrt{", with: "√(")
        result = result.replacingOccurrences(of: "\\sum", with: "∑")
        result = result.replacingOccurrences(of: "\\int", with: "∫")
        result = result.replacingOccurrences(of: "\\alpha", with: "α")
        result = result.replacingOccurrences(of: "\\beta", with: "β")
        result = result.replacingOccurrences(of: "\\gamma", with: "γ")
        result = result.replacingOccurrences(of: "\\pi", with: "π")
        result = result.replacingOccurrences(of: "\\infty", with: "∞")
        result = result.replacingOccurrences(of: "^{", with: "^")
        result = result.replacingOccurrences(of: "_{", with: "_")
        result = result.replacingOccurrences(of: "}", with: "")
        result = result.replacingOccurrences(of: "{", with: "")
        result = result.replacingOccurrences(of: "\\", with: "")
        return result
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - Advanced Text Formatting

/// 段落邊框
struct ParagraphBorder {
    var top: ParagraphBorderStyle?
    var bottom: ParagraphBorderStyle?
    var left: ParagraphBorderStyle?
    var right: ParagraphBorderStyle?
    var between: ParagraphBorderStyle?  // 段落之間的邊框

    init(
        top: ParagraphBorderStyle? = nil,
        bottom: ParagraphBorderStyle? = nil,
        left: ParagraphBorderStyle? = nil,
        right: ParagraphBorderStyle? = nil,
        between: ParagraphBorderStyle? = nil
    ) {
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
        self.between = between
    }

    /// 便利方法：建立四邊相同的邊框
    static func all(_ style: ParagraphBorderStyle) -> ParagraphBorder {
        ParagraphBorder(top: style, bottom: style, left: style, right: style)
    }
}

/// 邊框樣式
struct ParagraphBorderStyle {
    var type: ParagraphBorderType
    var color: String       // 十六進位顏色碼
    var size: Int           // 邊框寬度 (1/8 點)
    var space: Int          // 與文字間距 (點)

    init(type: ParagraphBorderType = .single, color: String = "000000", size: Int = 4, space: Int = 1) {
        self.type = type
        self.color = color
        self.size = size
        self.space = space
    }
}

/// 邊框類型（段落專用）
enum ParagraphBorderType: String, Codable {
    case none = "none"
    case single = "single"
    case thick = "thick"
    case double = "double"
    case dotted = "dotted"
    case dashed = "dashed"
    case dashDotStroked = "dashDotStroked"
    case threeDEmboss = "threeDEmboss"
    case threeDEngrave = "threeDEngrave"
    case wave = "wave"
}

extension ParagraphBorder {
    /// 產生段落邊框的 XML
    func toXML() -> String {
        var xml = "<w:pBdr>"

        if let top = top {
            xml += "<w:top w:val=\"\(top.type.rawValue)\" w:sz=\"\(top.size)\" w:space=\"\(top.space)\" w:color=\"\(top.color)\"/>"
        }
        if let bottom = bottom {
            xml += "<w:bottom w:val=\"\(bottom.type.rawValue)\" w:sz=\"\(bottom.size)\" w:space=\"\(bottom.space)\" w:color=\"\(bottom.color)\"/>"
        }
        if let left = left {
            xml += "<w:left w:val=\"\(left.type.rawValue)\" w:sz=\"\(left.size)\" w:space=\"\(left.space)\" w:color=\"\(left.color)\"/>"
        }
        if let right = right {
            xml += "<w:right w:val=\"\(right.type.rawValue)\" w:sz=\"\(right.size)\" w:space=\"\(right.space)\" w:color=\"\(right.color)\"/>"
        }
        if let between = between {
            xml += "<w:between w:val=\"\(between.type.rawValue)\" w:sz=\"\(between.size)\" w:space=\"\(between.space)\" w:color=\"\(between.color)\"/>"
        }

        xml += "</w:pBdr>"
        return xml
    }
}

/// 段落底色（使用 Table.swift 中的 CellShading 結構）
typealias ParagraphShading = CellShading

/// 字元間距設定
struct CharacterSpacing {
    var spacing: Int?       // 字元間距 (1/20 點，正值增加，負值減少)
    var position: Int?      // 位置調整（上升/下降）
    var kern: Int?          // 字距調整起始點數

    init(spacing: Int? = nil, position: Int? = nil, kern: Int? = nil) {
        self.spacing = spacing
        self.position = position
        self.kern = kern
    }
}

extension CharacterSpacing {
    /// 產生字元間距的 XML（在 rPr 內使用）
    func toXML() -> String {
        var xml = ""

        if let spacing = spacing {
            xml += "<w:spacing w:val=\"\(spacing)\"/>"
        }
        if let position = position {
            xml += "<w:position w:val=\"\(position)\"/>"
        }
        if let kern = kern {
            xml += "<w:kern w:val=\"\(kern)\"/>"
        }

        return xml
    }
}

/// 文字效果
enum TextEffect: String, Codable {
    case blinkBackground = "blinkBackground"
    case lights = "lights"
    case antsBlack = "antsBlack"
    case antsRed = "antsRed"
    case shimmer = "shimmer"
    case sparkle = "sparkle"
    case none = "none"
}

extension TextEffect {
    /// 產生文字效果的 XML（在 rPr 內使用）
    func toXML() -> String {
        if self == .none {
            return ""
        }
        return "<w:effect w:val=\"\(rawValue)\"/>"
    }
}
