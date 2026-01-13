import Foundation

/// 段落 (Paragraph) - Word 文件的基本結構單元
struct Paragraph {
    var runs: [Run]
    var properties: ParagraphProperties
    var hasPageBreak: Bool = false      // 是否為分頁符段落
    var bookmarks: [Bookmark] = []      // 段落內的書籤
    var hyperlinks: [Hyperlink] = []    // 段落內的超連結
    var commentIds: [Int] = []          // 段落關聯的註解 ID
    var footnoteIds: [Int] = []         // 段落內的腳註 ID
    var endnoteIds: [Int] = []          // 段落內的尾註 ID

    init(runs: [Run] = [], properties: ParagraphProperties = ParagraphProperties()) {
        self.runs = runs
        self.properties = properties
    }

    /// 便利初始化器：直接用文字建立段落
    init(text: String, properties: ParagraphProperties = ParagraphProperties()) {
        self.runs = [Run(text: text)]
        self.properties = properties
    }

    /// 取得段落純文字
    func getText() -> String {
        var text = runs.map { $0.text }.joined()
        // 加入超連結文字
        for hyperlink in hyperlinks {
            text += hyperlink.text
        }
        return text
    }
}

// MARK: - Paragraph Properties

/// 段落格式屬性
struct ParagraphProperties {
    var alignment: Alignment?
    var spacing: Spacing?
    var indentation: Indentation?
    var style: String?                  // 樣式名稱 (e.g., "Heading1")
    var numbering: NumberingInfo?       // 編號/項目符號
    var keepNext: Bool = false          // 與下段同頁
    var keepLines: Bool = false         // 段落不分頁
    var pageBreakBefore: Bool = false   // 段落前分頁
    var sectionBreak: SectionBreakType? // 分節符類型
    var border: ParagraphBorder?        // 段落邊框
    var shading: ParagraphShading?      // 段落底色

    init() {}

    /// 合併格式（覆蓋非 nil 值）
    mutating func merge(with other: ParagraphProperties) {
        if let alignment = other.alignment { self.alignment = alignment }
        if let spacing = other.spacing { self.spacing = spacing }
        if let indentation = other.indentation { self.indentation = indentation }
        if let style = other.style { self.style = style }
        if let numbering = other.numbering { self.numbering = numbering }
        if other.keepNext { self.keepNext = true }
        if other.keepLines { self.keepLines = true }
        if other.pageBreakBefore { self.pageBreakBefore = true }
        if let sectionBreak = other.sectionBreak { self.sectionBreak = sectionBreak }
        if let border = other.border { self.border = border }
        if let shading = other.shading { self.shading = shading }
    }
}

// MARK: - Supporting Types

/// 對齊方式
enum Alignment: String, Codable {
    case left = "left"
    case center = "center"
    case right = "right"
    case both = "both"      // 左右對齊（兩端對齊）
    case distribute = "distribute"  // 分散對齊
}

/// 段落對齊（Alignment 的別名）
typealias ParagraphAlignment = Alignment

/// 段落間距
struct Spacing {
    var before: Int?        // 段前間距 (1/20 點，twips)
    var after: Int?         // 段後間距 (1/20 點)
    var line: Int?          // 行高 (1/240 點 或 百分比)
    var lineRule: LineRule? // 行高規則

    init(before: Int? = nil, after: Int? = nil, line: Int? = nil, lineRule: LineRule? = nil) {
        self.before = before
        self.after = after
        self.line = line
        self.lineRule = lineRule
    }

    /// 便利方法：建立點數間距
    static func points(before: Double? = nil, after: Double? = nil, lineSpacing: Double? = nil) -> Spacing {
        var spacing = Spacing()
        if let before = before {
            spacing.before = Int(before * 20)  // 轉換為 twips
        }
        if let after = after {
            spacing.after = Int(after * 20)
        }
        if let lineSpacing = lineSpacing {
            spacing.line = Int(lineSpacing * 240)  // 固定行高
            spacing.lineRule = .exact
        }
        return spacing
    }
}

/// 行高規則
enum LineRule: String, Codable {
    case auto = "auto"          // 單行/1.5行/雙行
    case exact = "exact"        // 固定行高
    case atLeast = "atLeast"    // 最小行高
}

/// 縮排
struct Indentation {
    var left: Int?          // 左縮排 (twips)
    var right: Int?         // 右縮排 (twips)
    var firstLine: Int?     // 首行縮排 (twips)
    var hanging: Int?       // 凸排 (twips)

    init(left: Int? = nil, right: Int? = nil, firstLine: Int? = nil, hanging: Int? = nil) {
        self.left = left
        self.right = right
        self.firstLine = firstLine
        self.hanging = hanging
    }

    /// 便利方法：建立字元縮排（假設 1 字元 = 240 twips）
    static func characters(left: Int? = nil, right: Int? = nil, firstLine: Int? = nil) -> Indentation {
        var indent = Indentation()
        if let left = left {
            indent.left = left * 240
        }
        if let right = right {
            indent.right = right * 240
        }
        if let firstLine = firstLine {
            indent.firstLine = firstLine * 240
        }
        return indent
    }
}

/// 編號資訊
struct NumberingInfo {
    var numId: Int          // 編號定義 ID
    var level: Int          // 編號層級 (0-8)

    init(numId: Int, level: Int = 0) {
        self.numId = numId
        self.level = level
    }
}

// MARK: - XML 生成

extension Paragraph {
    /// 轉換為 OOXML XML 字串
    func toXML() -> String {
        var xml = "<w:p>"

        // Paragraph Properties
        let propsXML = properties.toXML()
        if !propsXML.isEmpty || properties.sectionBreak != nil {
            xml += "<w:pPr>\(propsXML)"

            // 分節符放在段落屬性中
            if let sectionBreak = properties.sectionBreak {
                xml += "<w:sectPr><w:type w:val=\"\(sectionBreak.rawValue)\"/></w:sectPr>"
            }

            xml += "</w:pPr>"
        }

        // 分頁符
        if hasPageBreak {
            xml += "<w:r><w:br w:type=\"page\"/></w:r>"
        }

        // 書籤開始標記
        for bookmark in bookmarks {
            xml += bookmark.toBookmarkStartXML()
        }

        // 註解範圍開始標記
        for commentId in commentIds {
            xml += "<w:commentRangeStart w:id=\"\(commentId)\"/>"
        }

        // Runs
        for run in runs {
            xml += run.toXML()
        }

        // 超連結
        for hyperlink in hyperlinks {
            xml += hyperlink.toXML()
        }

        // 註解範圍結束標記和參照
        for commentId in commentIds {
            xml += "<w:commentRangeEnd w:id=\"\(commentId)\"/>"
            xml += "<w:r><w:commentReference w:id=\"\(commentId)\"/></w:r>"
        }

        // 腳註參照
        for footnoteId in footnoteIds {
            xml += "<w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:footnoteReference w:id=\"\(footnoteId)\"/></w:r>"
        }

        // 尾註參照
        for endnoteId in endnoteIds {
            xml += "<w:r><w:rPr><w:rStyle w:val=\"EndnoteReference\"/></w:rPr><w:endnoteReference w:id=\"\(endnoteId)\"/></w:r>"
        }

        // 書籤結束標記
        for bookmark in bookmarks {
            xml += bookmark.toBookmarkEndXML()
        }

        xml += "</w:p>"
        return xml
    }
}

extension ParagraphProperties {
    /// 轉換為 OOXML XML 字串
    func toXML() -> String {
        var parts: [String] = []

        // 樣式
        if let style = style {
            parts.append("<w:pStyle w:val=\"\(style)\"/>")
        }

        // 編號
        if let numbering = numbering {
            parts.append("<w:numPr>")
            parts.append("<w:ilvl w:val=\"\(numbering.level)\"/>")
            parts.append("<w:numId w:val=\"\(numbering.numId)\"/>")
            parts.append("</w:numPr>")
        }

        // 對齊
        if let alignment = alignment {
            parts.append("<w:jc w:val=\"\(alignment.rawValue)\"/>")
        }

        // 間距
        if let spacing = spacing {
            var attrs: [String] = []
            if let before = spacing.before {
                attrs.append("w:before=\"\(before)\"")
            }
            if let after = spacing.after {
                attrs.append("w:after=\"\(after)\"")
            }
            if let line = spacing.line {
                attrs.append("w:line=\"\(line)\"")
            }
            if let lineRule = spacing.lineRule {
                attrs.append("w:lineRule=\"\(lineRule.rawValue)\"")
            }
            if !attrs.isEmpty {
                parts.append("<w:spacing \(attrs.joined(separator: " "))/>" )
            }
        }

        // 縮排
        if let indentation = indentation {
            var attrs: [String] = []
            if let left = indentation.left {
                attrs.append("w:left=\"\(left)\"")
            }
            if let right = indentation.right {
                attrs.append("w:right=\"\(right)\"")
            }
            if let firstLine = indentation.firstLine {
                attrs.append("w:firstLine=\"\(firstLine)\"")
            }
            if let hanging = indentation.hanging {
                attrs.append("w:hanging=\"\(hanging)\"")
            }
            if !attrs.isEmpty {
                parts.append("<w:ind \(attrs.joined(separator: " "))/>" )
            }
        }

        // 分頁控制
        if keepNext {
            parts.append("<w:keepNext/>")
        }
        if keepLines {
            parts.append("<w:keepLines/>")
        }
        if pageBreakBefore {
            parts.append("<w:pageBreakBefore/>")
        }

        // 段落邊框
        if let border = border {
            parts.append(border.toXML())
        }

        // 段落底色
        if let shading = shading {
            parts.append(shading.toXML())
        }

        return parts.joined()
    }
}
