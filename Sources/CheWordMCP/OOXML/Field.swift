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

// MARK: - Field Code Base

/// 通用欄位代碼基礎
protocol FieldCode {
    var fieldInstruction: String { get }
    var cachedResult: String? { get }
}

extension FieldCode {
    /// 產生欄位 XML
    func toFieldXML() -> String {
        var xml = ""
        xml += "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
        xml += "<w:r><w:instrText xml:space=\"preserve\"> \(fieldInstruction) </w:instrText></w:r>"
        xml += "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        xml += "<w:r><w:t>\(escapeXMLForField(cachedResult ?? ""))</w:t></w:r>"
        xml += "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
        return xml
    }

    private func escapeXMLForField(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - IF Field (條件判斷欄位)

/// IF 條件判斷欄位
/// 語法: IF expression operator expression "trueText" "falseText"
struct IFField: FieldCode {
    var leftOperand: String         // 左運算元（可以是欄位參照或值）
    var comparisonOperator: ComparisonOperator
    var rightOperand: String        // 右運算元
    var trueText: String            // 條件為真時顯示
    var falseText: String           // 條件為假時顯示
    var cachedResult: String?

    /// 比較運算子
    enum ComparisonOperator: String {
        case equal = "="
        case notEqual = "<>"
        case lessThan = "<"
        case greaterThan = ">"
        case lessThanOrEqual = "<="
        case greaterThanOrEqual = ">="
    }

    init(
        leftOperand: String,
        comparisonOperator: ComparisonOperator,
        rightOperand: String,
        trueText: String,
        falseText: String,
        cachedResult: String? = nil
    ) {
        self.leftOperand = leftOperand
        self.comparisonOperator = comparisonOperator
        self.rightOperand = rightOperand
        self.trueText = trueText
        self.falseText = falseText
        self.cachedResult = cachedResult
    }

    var fieldInstruction: String {
        // IF "leftOperand" operator "rightOperand" "trueText" "falseText"
        return "IF \"\(leftOperand)\" \(comparisonOperator.rawValue) \"\(rightOperand)\" \"\(trueText)\" \"\(falseText)\""
    }

    /// 便利建構：檢查欄位是否為空
    static func ifEmpty(fieldName: String, trueText: String, falseText: String) -> IFField {
        return IFField(
            leftOperand: "{ MERGEFIELD \(fieldName) }",
            comparisonOperator: .equal,
            rightOperand: "",
            trueText: trueText,
            falseText: falseText
        )
    }

    /// 便利建構：檢查欄位是否等於某值
    static func ifEquals(fieldName: String, value: String, trueText: String, falseText: String) -> IFField {
        return IFField(
            leftOperand: "{ MERGEFIELD \(fieldName) }",
            comparisonOperator: .equal,
            rightOperand: value,
            trueText: trueText,
            falseText: falseText
        )
    }
}

// MARK: - Calculation Field (計算欄位)

/// 計算欄位
/// 語法: = expression [bookmark] [\# "numeric-picture"]
struct CalculationField: FieldCode {
    var expression: String          // 數學表達式
    var bookmark: String?           // 可選書籤名稱（用於儲存結果）
    var numberFormat: String?       // 數字格式（如 "#,##0.00"）
    var cachedResult: String?

    init(expression: String, bookmark: String? = nil, numberFormat: String? = nil, cachedResult: String? = nil) {
        self.expression = expression
        self.bookmark = bookmark
        self.numberFormat = numberFormat
        self.cachedResult = cachedResult
    }

    var fieldInstruction: String {
        var instruction = "= \(expression)"
        if let bookmark = bookmark {
            instruction += " \(bookmark)"
        }
        if let format = numberFormat {
            instruction += " \\# \"\(format)\""
        }
        return instruction
    }

    /// 基本運算子
    enum Operator: String {
        case add = "+"
        case subtract = "-"
        case multiply = "*"
        case divide = "/"
        case percent = "%"
        case power = "^"
    }

    /// 便利建構：加總表格儲存格
    static func sum(above: Bool = true, bookmark: String? = nil, format: String? = "#,##0") -> CalculationField {
        return CalculationField(
            expression: above ? "SUM(ABOVE)" : "SUM(LEFT)",
            bookmark: bookmark,
            numberFormat: format
        )
    }

    /// 便利建構：平均值
    static func average(above: Bool = true, bookmark: String? = nil, format: String? = "#,##0.00") -> CalculationField {
        return CalculationField(
            expression: above ? "AVERAGE(ABOVE)" : "AVERAGE(LEFT)",
            bookmark: bookmark,
            numberFormat: format
        )
    }

    /// 便利建構：計數
    static func count(above: Bool = true, bookmark: String? = nil) -> CalculationField {
        return CalculationField(
            expression: above ? "COUNT(ABOVE)" : "COUNT(LEFT)",
            bookmark: bookmark,
            numberFormat: "#,##0"
        )
    }

    /// 便利建構：最大值
    static func max(above: Bool = true, bookmark: String? = nil, format: String? = "#,##0") -> CalculationField {
        return CalculationField(
            expression: above ? "MAX(ABOVE)" : "MAX(LEFT)",
            bookmark: bookmark,
            numberFormat: format
        )
    }

    /// 便利建構：最小值
    static func min(above: Bool = true, bookmark: String? = nil, format: String? = "#,##0") -> CalculationField {
        return CalculationField(
            expression: above ? "MIN(ABOVE)" : "MIN(LEFT)",
            bookmark: bookmark,
            numberFormat: format
        )
    }

    /// 便利建構：書籤相乘
    static func multiply(bookmark1: String, bookmark2: String, format: String? = "#,##0") -> CalculationField {
        return CalculationField(
            expression: "\(bookmark1)*\(bookmark2)",
            numberFormat: format
        )
    }
}

// MARK: - Date/Time Fields (日期時間欄位)

/// 日期時間欄位類型
enum DateTimeFieldType: String {
    case date = "DATE"              // 目前日期
    case time = "TIME"              // 目前時間
    case createDate = "CREATEDATE"  // 文件建立日期
    case saveDate = "SAVEDATE"      // 最後儲存日期
    case printDate = "PRINTDATE"    // 最後列印日期
    case editTime = "EDITTIME"      // 總編輯時間（分鐘）
}

/// 日期時間欄位
struct DateTimeField: FieldCode {
    var type: DateTimeFieldType
    var dateFormat: String?         // 日期格式（如 "yyyy/MM/dd"）
    var useLastUsedFormat: Bool     // 使用上次使用的格式
    var cachedResult: String?

    /// 常用日期格式
    enum DateFormat: String {
        case shortDate = "yyyy/M/d"
        case longDate = "yyyy年M月d日"
        case shortDateTime = "yyyy/M/d H:mm"
        case longDateTime = "yyyy年M月d日 H:mm:ss"
        case timeOnly = "H:mm:ss"
        case monthYear = "yyyy年M月"
        case dayMonthYear = "d/M/yyyy"
        case iso8601 = "yyyy-MM-dd"
        case custom = ""
    }

    init(
        type: DateTimeFieldType = .date,
        dateFormat: String? = nil,
        useLastUsedFormat: Bool = false,
        cachedResult: String? = nil
    ) {
        self.type = type
        self.dateFormat = dateFormat
        self.useLastUsedFormat = useLastUsedFormat
        self.cachedResult = cachedResult
    }

    var fieldInstruction: String {
        var instruction = type.rawValue
        if useLastUsedFormat {
            instruction += " \\l"
        } else if let format = dateFormat {
            instruction += " \\@ \"\(format)\""
        }
        return instruction
    }

    /// 便利建構：今天日期
    static func today(format: DateFormat = .shortDate) -> DateTimeField {
        return DateTimeField(type: .date, dateFormat: format.rawValue)
    }

    /// 便利建構：目前時間
    static func now(includeDate: Bool = false) -> DateTimeField {
        if includeDate {
            return DateTimeField(type: .date, dateFormat: DateFormat.shortDateTime.rawValue)
        } else {
            return DateTimeField(type: .time, dateFormat: DateFormat.timeOnly.rawValue)
        }
    }

    /// 便利建構：文件建立日期
    static func created(format: DateFormat = .longDate) -> DateTimeField {
        return DateTimeField(type: .createDate, dateFormat: format.rawValue)
    }

    /// 便利建構：最後修改日期
    static func lastSaved(format: DateFormat = .longDateTime) -> DateTimeField {
        return DateTimeField(type: .saveDate, dateFormat: format.rawValue)
    }
}

// MARK: - Document Information Fields (文件資訊欄位)

/// 文件資訊欄位類型
enum DocumentInfoFieldType: String {
    // 頁面資訊
    case page = "PAGE"              // 目前頁碼
    case numPages = "NUMPAGES"      // 總頁數
    case sectionPages = "SECTIONPAGES"  // 本節頁數

    // 文件屬性
    case fileName = "FILENAME"       // 檔案名稱
    case fileSize = "FILESIZE"       // 檔案大小
    case author = "AUTHOR"           // 作者
    case title = "TITLE"             // 標題
    case subject = "SUBJECT"         // 主旨
    case keywords = "KEYWORDS"       // 關鍵字
    case comments = "COMMENTS"       // 註解
    case lastSavedBy = "LASTSAVEDBY" // 最後儲存者
    case revNum = "REVNUM"           // 修訂版本號

    // 統計資訊
    case numChars = "NUMCHARS"       // 字元數
    case numWords = "NUMWORDS"       // 字數
}

/// 文件資訊欄位
struct DocumentInfoField: FieldCode {
    var type: DocumentInfoFieldType
    var includePath: Bool           // 是否包含路徑（用於 FILENAME）
    var format: String?             // 格式（用於數字類型）
    var cachedResult: String?

    init(
        type: DocumentInfoFieldType,
        includePath: Bool = false,
        format: String? = nil,
        cachedResult: String? = nil
    ) {
        self.type = type
        self.includePath = includePath
        self.format = format
        self.cachedResult = cachedResult
    }

    var fieldInstruction: String {
        var instruction = type.rawValue
        if type == .fileName && includePath {
            instruction += " \\p"
        }
        if let format = format {
            instruction += " \\# \"\(format)\""
        }
        return instruction
    }

    /// 便利建構：頁碼
    static func pageNumber(format: String? = nil) -> DocumentInfoField {
        return DocumentInfoField(type: .page, format: format)
    }

    /// 便利建構：頁碼/總頁數 格式
    static func pageOfTotal() -> String {
        let pageField = DocumentInfoField(type: .page)
        let totalField = DocumentInfoField(type: .numPages)
        return pageField.toFieldXML() + "<w:r><w:t>/</w:t></w:r>" + totalField.toFieldXML()
    }

    /// 便利建構：檔案名稱
    static func fileName(withPath: Bool = false) -> DocumentInfoField {
        return DocumentInfoField(type: .fileName, includePath: withPath)
    }

    /// 便利建構：作者
    static func author() -> DocumentInfoField {
        return DocumentInfoField(type: .author)
    }

    /// 便利建構：字數
    static func wordCount(format: String? = "#,##0") -> DocumentInfoField {
        return DocumentInfoField(type: .numWords, format: format)
    }
}

// MARK: - Reference Fields (參照欄位)

/// 參照欄位類型
enum ReferenceFieldType: String {
    case ref = "REF"                // 書籤參照
    case pageRef = "PAGEREF"        // 頁碼參照
    case noteRef = "NOTEREF"        // 註腳/尾註參照
}

/// 參照欄位
struct ReferenceField: FieldCode {
    var type: ReferenceFieldType
    var bookmarkName: String        // 書籤名稱
    var includeAboveBelow: Bool     // 是否包含「上方」/「下方」文字
    var createHyperlink: Bool       // 是否建立超連結
    var cachedResult: String?

    init(
        type: ReferenceFieldType = .ref,
        bookmarkName: String,
        includeAboveBelow: Bool = false,
        createHyperlink: Bool = false,
        cachedResult: String? = nil
    ) {
        self.type = type
        self.bookmarkName = bookmarkName
        self.includeAboveBelow = includeAboveBelow
        self.createHyperlink = createHyperlink
        self.cachedResult = cachedResult
    }

    var fieldInstruction: String {
        var instruction = "\(type.rawValue) \(bookmarkName)"
        if includeAboveBelow {
            instruction += " \\p"
        }
        if createHyperlink {
            instruction += " \\h"
        }
        return instruction
    }

    /// 便利建構：書籤參照
    static func bookmark(_ name: String, hyperlink: Bool = true) -> ReferenceField {
        return ReferenceField(type: .ref, bookmarkName: name, createHyperlink: hyperlink)
    }

    /// 便利建構：頁碼參照
    static func pageOf(_ bookmarkName: String, hyperlink: Bool = true) -> ReferenceField {
        return ReferenceField(type: .pageRef, bookmarkName: bookmarkName, createHyperlink: hyperlink)
    }
}

// MARK: - Sequence Field (序列欄位)

/// 序列欄位（自動編號）
struct SequenceField: FieldCode {
    var identifier: String          // 序列識別符（如 "Figure", "Table"）
    var format: SequenceFormat      // 編號格式
    var resetLevel: Int?            // 重設層級（對應標題層級）
    var hideResult: Bool            // 是否隱藏結果
    var cachedResult: String?

    /// 序列編號格式
    enum SequenceFormat: String {
        case arabic = ""            // 1, 2, 3
        case alphabetic = "\\* ALPHABETIC"  // A, B, C
        case lowerAlphabetic = "\\* alphabetic"  // a, b, c
        case roman = "\\* ROMAN"    // I, II, III
        case lowerRoman = "\\* roman"  // i, ii, iii
    }

    init(
        identifier: String,
        format: SequenceFormat = .arabic,
        resetLevel: Int? = nil,
        hideResult: Bool = false,
        cachedResult: String? = nil
    ) {
        self.identifier = identifier
        self.format = format
        self.resetLevel = resetLevel
        self.hideResult = hideResult
        self.cachedResult = cachedResult
    }

    var fieldInstruction: String {
        var instruction = "SEQ \(identifier)"
        if !format.rawValue.isEmpty {
            instruction += " \(format.rawValue)"
        }
        if let level = resetLevel {
            instruction += " \\s \(level)"
        }
        if hideResult {
            instruction += " \\h"
        }
        return instruction
    }

    /// 便利建構：圖片編號
    static func figure(resetOnHeading: Int? = 1) -> SequenceField {
        return SequenceField(identifier: "Figure", resetLevel: resetOnHeading)
    }

    /// 便利建構：表格編號
    static func table(resetOnHeading: Int? = 1) -> SequenceField {
        return SequenceField(identifier: "Table", resetLevel: resetOnHeading)
    }

    /// 便利建構：章節編號
    static func chapter() -> SequenceField {
        return SequenceField(identifier: "Chapter")
    }
}

// MARK: - Mail Merge Fields (合併列印欄位)

/// 合併列印欄位
struct MergeField: FieldCode {
    var fieldName: String           // 欄位名稱
    var textBefore: String?         // 前置文字（僅當欄位非空時顯示）
    var textAfter: String?          // 後置文字（僅當欄位非空時顯示）
    var isMapped: Bool              // 是否為對應欄位
    var verticalFormatting: Bool    // 是否垂直格式化
    var cachedResult: String?

    init(
        fieldName: String,
        textBefore: String? = nil,
        textAfter: String? = nil,
        isMapped: Bool = false,
        verticalFormatting: Bool = false,
        cachedResult: String? = nil
    ) {
        self.fieldName = fieldName
        self.textBefore = textBefore
        self.textAfter = textAfter
        self.isMapped = isMapped
        self.verticalFormatting = verticalFormatting
        self.cachedResult = cachedResult
    }

    var fieldInstruction: String {
        var instruction = "MERGEFIELD \(fieldName)"
        if let before = textBefore {
            instruction += " \\b \"\(before)\""
        }
        if let after = textAfter {
            instruction += " \\f \"\(after)\""
        }
        if isMapped {
            instruction += " \\m"
        }
        if verticalFormatting {
            instruction += " \\v"
        }
        return instruction
    }

    /// 便利建構：簡單欄位
    static func simple(_ fieldName: String) -> MergeField {
        return MergeField(fieldName: fieldName)
    }

    /// 便利建構：帶前後文字的欄位
    static func withContext(_ fieldName: String, before: String? = nil, after: String? = nil) -> MergeField {
        return MergeField(fieldName: fieldName, textBefore: before, textAfter: after)
    }
}

// MARK: - Nested Fields (巢狀欄位)

/// 巢狀欄位（欄位內包含欄位）
struct NestedField {
    var outerInstruction: String
    var innerFields: [any FieldCode]
    var cachedResult: String?

    /// 產生巢狀欄位 XML（使用特殊分隔符標記）
    func toNestedXML() -> String {
        var xml = ""
        xml += "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"

        // 外部指令的開頭部分
        let parts = outerInstruction.split(separator: "{", maxSplits: 1)
        if let firstPart = parts.first {
            xml += "<w:r><w:instrText xml:space=\"preserve\"> \(String(firstPart))</w:instrText></w:r>"
        }

        // 插入內部欄位
        for innerField in innerFields {
            xml += innerField.toFieldXML()
        }

        // 外部指令的結尾部分（如果有）
        if parts.count > 1 {
            xml += "<w:r><w:instrText xml:space=\"preserve\">\(String(parts[1]))</w:instrText></w:r>"
        }

        xml += "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        xml += "<w:r><w:t>\(cachedResult ?? "")</w:t></w:r>"
        xml += "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
        return xml
    }
}

// MARK: - Field Collection

/// 欄位集合（用於管理文件中的所有欄位）
struct FieldCollection {
    var fields: [any FieldCode] = []

    /// 新增欄位
    mutating func addField(_ field: any FieldCode) {
        fields.append(field)
    }

    /// 新增 IF 欄位
    mutating func addIF(_ ifField: IFField) {
        fields.append(ifField)
    }

    /// 新增計算欄位
    mutating func addCalculation(_ calcField: CalculationField) {
        fields.append(calcField)
    }

    /// 新增日期時間欄位
    mutating func addDateTime(_ dateField: DateTimeField) {
        fields.append(dateField)
    }

    /// 新增文件資訊欄位
    mutating func addDocumentInfo(_ infoField: DocumentInfoField) {
        fields.append(infoField)
    }

    /// 新增參照欄位
    mutating func addReference(_ refField: ReferenceField) {
        fields.append(refField)
    }

    /// 新增序列欄位
    mutating func addSequence(_ seqField: SequenceField) {
        fields.append(seqField)
    }

    /// 新增合併欄位
    mutating func addMergeField(_ mergeField: MergeField) {
        fields.append(mergeField)
    }
}

// MARK: - Field Errors

/// 欄位錯誤
enum FieldError: Error, LocalizedError {
    case invalidExpression(String)
    case bookmarkNotFound(String)
    case invalidFormat(String)
    case nestedFieldTooDeep
    case unsupportedFieldType(String)

    var errorDescription: String? {
        switch self {
        case .invalidExpression(let expr):
            return "Invalid field expression: \(expr)"
        case .bookmarkNotFound(let name):
            return "Bookmark not found: \(name)"
        case .invalidFormat(let format):
            return "Invalid format string: \(format)"
        case .nestedFieldTooDeep:
            return "Nested fields exceed maximum depth"
        case .unsupportedFieldType(let type):
            return "Unsupported field type: \(type)"
        }
    }
}

// MARK: - Structured Document Tags (SDT) - Repeating Sections

/// SDT 類型
enum SDTType: String {
    case richText = "richText"              // 富文本
    case plainText = "text"                 // 純文字
    case picture = "picture"                // 圖片
    case date = "date"                      // 日期
    case dropDownList = "dropDownList"      // 下拉選單
    case comboBox = "comboBox"              // 組合方塊
    case checkbox = "checkbox"              // 核取方塊
    case bibliography = "bibliography"       // 書目
    case citation = "citation"              // 引用
    case group = "group"                    // 群組
    case repeatingSection = "repeatingSection"  // 重複區段
    case repeatingSectionItem = "repeatingSectionItem"  // 重複區段項目
}

/// SDT 鎖定類型
enum SDTLockType: String {
    case unlocked = ""                      // 未鎖定
    case sdtLocked = "sdtLocked"            // SDT 鎖定（無法刪除 SDT）
    case contentLocked = "contentLocked"    // 內容鎖定（無法編輯內容）
    case sdtContentLocked = "sdtContentLocked"  // 完全鎖定
}

/// 結構化文件標籤（SDT）基礎
struct StructuredDocumentTag {
    var id: Int?                            // SDT 唯一 ID
    var tag: String?                        // 標籤（用於識別）
    var alias: String?                      // 顯示名稱
    var type: SDTType                       // SDT 類型
    var lockType: SDTLockType               // 鎖定類型
    var placeholder: String?                // 佔位符提示文字
    var isTemporary: Bool                   // 是否為暫時（編輯後移除 SDT）
    var showAsHidden: Bool                  // 是否以隱藏框顯示

    init(
        id: Int? = nil,
        tag: String? = nil,
        alias: String? = nil,
        type: SDTType = .richText,
        lockType: SDTLockType = .unlocked,
        placeholder: String? = nil,
        isTemporary: Bool = false,
        showAsHidden: Bool = false
    ) {
        self.id = id
        self.tag = tag
        self.alias = alias
        self.type = type
        self.lockType = lockType
        self.placeholder = placeholder
        self.isTemporary = isTemporary
        self.showAsHidden = showAsHidden
    }

    /// 產生 SDT 屬性 XML
    func toSdtPrXML() -> String {
        var xml = "<w:sdtPr>"

        if let id = id {
            xml += "<w:id w:val=\"\(id)\"/>"
        }
        if let tag = tag {
            xml += "<w:tag w:val=\"\(escapeXML(tag))\"/>"
        }
        if let alias = alias {
            xml += "<w:alias w:val=\"\(escapeXML(alias))\"/>"
        }

        // 鎖定設定
        if lockType != .unlocked {
            xml += "<w:lock w:val=\"\(lockType.rawValue)\"/>"
        }

        // 佔位符
        if let placeholder = placeholder {
            xml += """
            <w:placeholder>
                <w:docPart w:val="\(escapeXML(placeholder))"/>
            </w:placeholder>
            <w:showingPlcHdr/>
            """
        }

        // 暫時標記
        if isTemporary {
            xml += "<w:temporary/>"
        }

        // 類型特定標記
        switch type {
        case .richText:
            break  // 預設，不需要額外標記
        case .plainText:
            xml += "<w:text/>"
        case .picture:
            xml += "<w:picture/>"
        case .date:
            xml += "<w:date/>"
        case .dropDownList:
            xml += "<w:dropDownList/>"
        case .comboBox:
            xml += "<w:comboBox/>"
        case .checkbox:
            xml += "<w14:checkbox xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"/>"
        case .bibliography:
            xml += "<w:bibliography/>"
        case .citation:
            xml += "<w:citation/>"
        case .group:
            xml += "<w:group/>"
        case .repeatingSection:
            xml += "<w15:repeatingSection xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"/>"
        case .repeatingSectionItem:
            xml += "<w15:repeatingSectionItem xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"/>"
        }

        xml += "</w:sdtPr>"
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

// MARK: - Repeating Section (重複區段)

/// 重複區段
struct RepeatingSection {
    var sdt: StructuredDocumentTag
    var items: [RepeatingSectionItem]       // 重複項目列表
    var allowInsertDeleteSections: Bool     // 允許插入/刪除區段
    var sectionTitle: String?               // 區段標題（顯示在 UI）

    init(
        tag: String? = nil,
        alias: String? = nil,
        items: [RepeatingSectionItem] = [],
        allowInsertDeleteSections: Bool = true,
        sectionTitle: String? = nil
    ) {
        self.sdt = StructuredDocumentTag(
            tag: tag,
            alias: alias,
            type: .repeatingSection
        )
        self.items = items
        self.allowInsertDeleteSections = allowInsertDeleteSections
        self.sectionTitle = sectionTitle
    }

    /// 新增項目
    mutating func addItem(_ item: RepeatingSectionItem) {
        items.append(item)
    }

    /// 新增空白項目
    mutating func addEmptyItem(content: String = "") {
        let item = RepeatingSectionItem(content: content)
        items.append(item)
    }

    /// 產生完整的重複區段 XML
    func toXML() -> String {
        var xml = "<w:sdt>"

        // SDT 屬性
        xml += "<w:sdtPr>"
        if let tag = sdt.tag {
            xml += "<w:tag w:val=\"\(escapeXML(tag))\"/>"
        }
        if let alias = sdt.alias {
            xml += "<w:alias w:val=\"\(escapeXML(alias))\"/>"
        }

        // 重複區段設定
        xml += "<w15:repeatingSection xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\">"
        if let title = sectionTitle {
            xml += "<w15:sectionTitle w15:val=\"\(escapeXML(title))\"/>"
        }
        if allowInsertDeleteSections {
            xml += "<w15:allowInsertDeleteSection w15:val=\"true\"/>"
        }
        xml += "</w15:repeatingSection>"
        xml += "</w:sdtPr>"

        // SDT 結束標籤屬性
        xml += "<w:sdtEndPr/>"

        // SDT 內容
        xml += "<w:sdtContent>"
        for item in items {
            xml += item.toXML()
        }
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

/// 重複區段項目
struct RepeatingSectionItem {
    var sdt: StructuredDocumentTag
    var content: String                     // 項目內容（可以是段落 XML）
    var paragraphs: [Paragraph]?            // 可選：使用 Paragraph 結構
    var tableRows: [TableRow]?              // 可選：用於表格行重複

    init(
        tag: String? = nil,
        content: String = "",
        paragraphs: [Paragraph]? = nil,
        tableRows: [TableRow]? = nil
    ) {
        self.sdt = StructuredDocumentTag(
            tag: tag,
            type: .repeatingSectionItem
        )
        self.content = content
        self.paragraphs = paragraphs
        self.tableRows = tableRows
    }

    /// 產生重複區段項目 XML
    func toXML() -> String {
        var xml = "<w:sdt>"

        // SDT 屬性
        xml += "<w:sdtPr>"
        if let tag = sdt.tag {
            xml += "<w:tag w:val=\"\(escapeXML(tag))\"/>"
        }
        xml += "<w15:repeatingSectionItem xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"/>"
        xml += "</w:sdtPr>"

        // SDT 內容
        xml += "<w:sdtContent>"

        // 優先使用 paragraphs
        if let paragraphs = paragraphs {
            for para in paragraphs {
                xml += para.toXML()
            }
        }
        // 然後檢查 tableRows
        else if let rows = tableRows {
            for row in rows {
                xml += row.toXML()
            }
        }
        // 最後使用簡單內容
        else if !content.isEmpty {
            xml += "<w:p><w:r><w:t>\(escapeXML(content))</w:t></w:r></w:p>"
        }

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

// MARK: - Content Control (內容控制項)

/// 內容控制項（通用 SDT 包裝）
struct ContentControl {
    var sdt: StructuredDocumentTag
    var content: String                     // 內容 XML 或純文字

    /// 建立富文本控制項
    static func richText(tag: String, alias: String, content: String, placeholder: String? = nil) -> ContentControl {
        return ContentControl(
            sdt: StructuredDocumentTag(
                tag: tag,
                alias: alias,
                type: .richText,
                placeholder: placeholder
            ),
            content: content
        )
    }

    /// 建立純文字控制項
    static func plainText(tag: String, alias: String, content: String, placeholder: String? = nil) -> ContentControl {
        return ContentControl(
            sdt: StructuredDocumentTag(
                tag: tag,
                alias: alias,
                type: .plainText,
                placeholder: placeholder
            ),
            content: content
        )
    }

    /// 建立日期控制項
    static func date(tag: String, alias: String, dateFormat: String = "yyyy/M/d", placeholder: String? = nil) -> ContentControl {
        return ContentControl(
            sdt: StructuredDocumentTag(
                tag: tag,
                alias: alias,
                type: .date,
                placeholder: placeholder
            ),
            content: ""
        )
    }

    /// 建立圖片控制項
    static func picture(tag: String, alias: String) -> ContentControl {
        return ContentControl(
            sdt: StructuredDocumentTag(
                tag: tag,
                alias: alias,
                type: .picture
            ),
            content: ""
        )
    }

    /// 產生內容控制項 XML
    func toXML() -> String {
        var xml = "<w:sdt>"
        xml += sdt.toSdtPrXML()
        xml += "<w:sdtContent>"
        if content.hasPrefix("<w:") {
            // 已經是 XML 格式
            xml += content
        } else if !content.isEmpty {
            // 純文字，包裝成段落
            xml += "<w:p><w:r><w:t>\(escapeXML(content))</w:t></w:r></w:p>"
        } else {
            // 空內容，插入空段落
            xml += "<w:p><w:r><w:t></w:t></w:r></w:p>"
        }
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

// MARK: - Data Binding (資料繫結)

/// 資料繫結（用於將 SDT 繫結到自訂 XML 部分）
struct DataBinding {
    var prefixMappings: String?             // XML 命名空間前綴對應
    var xpath: String                       // XPath 表達式
    var storeItemID: String?                // 自訂 XML 部分的 ID

    init(xpath: String, storeItemID: String? = nil, prefixMappings: String? = nil) {
        self.xpath = xpath
        self.storeItemID = storeItemID
        self.prefixMappings = prefixMappings
    }

    /// 產生資料繫結 XML
    func toXML() -> String {
        var attrs: [String] = []

        if let prefixMappings = prefixMappings {
            attrs.append("w:prefixMappings=\"\(escapeXML(prefixMappings))\"")
        }
        attrs.append("w:xpath=\"\(escapeXML(xpath))\"")
        if let storeItemID = storeItemID {
            attrs.append("w:storeItemID=\"\(escapeXML(storeItemID))\"")
        }

        return "<w:dataBinding \(attrs.joined(separator: " "))/>"
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

/// 帶資料繫結的內容控制項
struct BoundContentControl {
    var sdt: StructuredDocumentTag
    var dataBinding: DataBinding
    var content: String

    /// 產生帶繫結的控制項 XML
    func toXML() -> String {
        var xml = "<w:sdt>"

        // SDT 屬性
        xml += "<w:sdtPr>"
        if let id = sdt.id {
            xml += "<w:id w:val=\"\(id)\"/>"
        }
        if let tag = sdt.tag {
            xml += "<w:tag w:val=\"\(escapeXML(tag))\"/>"
        }
        if let alias = sdt.alias {
            xml += "<w:alias w:val=\"\(escapeXML(alias))\"/>"
        }

        // 類型
        switch sdt.type {
        case .plainText:
            xml += "<w:text/>"
        case .date:
            xml += "<w:date/>"
        default:
            break
        }

        // 資料繫結
        xml += dataBinding.toXML()

        xml += "</w:sdtPr>"

        // 內容
        xml += "<w:sdtContent>"
        if content.hasPrefix("<w:") {
            xml += content
        } else if !content.isEmpty {
            xml += "<w:r><w:t>\(escapeXML(content))</w:t></w:r>"
        }
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

// MARK: - SDT Collection

/// SDT 集合（用於管理文件中的所有內容控制項）
struct SDTCollection {
    var contentControls: [ContentControl] = []
    var repeatingSections: [RepeatingSection] = []
    var boundControls: [BoundContentControl] = []

    private var nextId: Int = 1

    /// 取得下一個 ID
    mutating func getNextId() -> Int {
        let id = nextId
        nextId += 1
        return id
    }

    /// 新增內容控制項
    mutating func addContentControl(_ control: ContentControl) {
        contentControls.append(control)
    }

    /// 新增重複區段
    mutating func addRepeatingSection(_ section: RepeatingSection) {
        repeatingSections.append(section)
    }

    /// 新增帶繫結的控制項
    mutating func addBoundControl(_ control: BoundContentControl) {
        boundControls.append(control)
    }

    /// 建立簡單的重複區段（用於列表項目）
    mutating func createSimpleRepeatingSection(tag: String, items: [String]) -> RepeatingSection {
        var section = RepeatingSection(tag: tag, alias: tag)
        for item in items {
            section.addEmptyItem(content: item)
        }
        repeatingSections.append(section)
        return section
    }
}

// MARK: - SDT Error

/// SDT 錯誤
enum SDTError: Error, LocalizedError {
    case invalidType(String)
    case contentNotAllowed(SDTType)
    case dataBindingFailed(String)
    case repeatingSectionEmpty
    case lockViolation

    var errorDescription: String? {
        switch self {
        case .invalidType(let type):
            return "Invalid SDT type: \(type)"
        case .contentNotAllowed(let type):
            return "Content not allowed for SDT type: \(type)"
        case .dataBindingFailed(let reason):
            return "Data binding failed: \(reason)"
        case .repeatingSectionEmpty:
            return "Repeating section must have at least one item"
        case .lockViolation:
            return "Cannot modify locked content control"
        }
    }
}
