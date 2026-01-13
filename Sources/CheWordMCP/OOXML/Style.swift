import Foundation

/// 樣式 (Style) - 定義段落和文字的預設格式
struct Style {
    var id: String              // 樣式 ID
    var name: String            // 顯示名稱
    var type: StyleType         // 樣式類型
    var basedOn: String?        // 基於的樣式
    var nextStyle: String?      // 下一段使用的樣式
    var isDefault: Bool         // 是否為預設樣式
    var isQuickStyle: Bool      // 是否顯示在快速樣式庫

    var paragraphProperties: ParagraphProperties?
    var runProperties: RunProperties?

    init(id: String,
         name: String,
         type: StyleType,
         basedOn: String? = nil,
         nextStyle: String? = nil,
         isDefault: Bool = false,
         isQuickStyle: Bool = true,
         paragraphProperties: ParagraphProperties? = nil,
         runProperties: RunProperties? = nil) {
        self.id = id
        self.name = name
        self.type = type
        self.basedOn = basedOn
        self.nextStyle = nextStyle
        self.isDefault = isDefault
        self.isQuickStyle = isQuickStyle
        self.paragraphProperties = paragraphProperties
        self.runProperties = runProperties
    }
}

// MARK: - Style Type

/// 樣式類型
enum StyleType: String, Codable {
    case paragraph = "paragraph"
    case character = "character"
    case table = "table"
    case numbering = "numbering"
}

// MARK: - Style Update

/// 樣式更新資料結構
struct StyleUpdate {
    var name: String?
    var basedOn: String?
    var nextStyle: String?
    var isQuickStyle: Bool?
    var paragraphProperties: ParagraphProperties?
    var runProperties: RunProperties?

    init(name: String? = nil,
         basedOn: String? = nil,
         nextStyle: String? = nil,
         isQuickStyle: Bool? = nil,
         paragraphProperties: ParagraphProperties? = nil,
         runProperties: RunProperties? = nil) {
        self.name = name
        self.basedOn = basedOn
        self.nextStyle = nextStyle
        self.isQuickStyle = isQuickStyle
        self.paragraphProperties = paragraphProperties
        self.runProperties = runProperties
    }
}

// MARK: - Default Styles

extension Style {
    /// 預設樣式集合
    static var defaultStyles: [Style] {
        return [
            normalStyle,
            heading1Style,
            heading2Style,
            heading3Style,
            titleStyle,
            subtitleStyle
        ]
    }

    /// Normal（內文）
    static var normalStyle: Style {
        var paraProps = ParagraphProperties()
        paraProps.spacing = Spacing(after: 200, line: 276, lineRule: .auto)

        var runProps = RunProperties()
        runProps.fontSize = 22      // 11pt
        runProps.fontName = "Calibri"

        return Style(
            id: "Normal",
            name: "Normal",
            type: .paragraph,
            isDefault: true,
            paragraphProperties: paraProps,
            runProperties: runProps
        )
    }

    /// 標題 1
    static var heading1Style: Style {
        var paraProps = ParagraphProperties()
        paraProps.spacing = Spacing(before: 240, after: 0)
        paraProps.keepNext = true
        paraProps.keepLines = true

        var runProps = RunProperties()
        runProps.fontSize = 32      // 16pt
        runProps.fontName = "Calibri Light"
        runProps.color = "2F5496"   // 深藍色
        runProps.bold = true

        return Style(
            id: "Heading1",
            name: "heading 1",
            type: .paragraph,
            basedOn: "Normal",
            nextStyle: "Normal",
            paragraphProperties: paraProps,
            runProperties: runProps
        )
    }

    /// 標題 2
    static var heading2Style: Style {
        var paraProps = ParagraphProperties()
        paraProps.spacing = Spacing(before: 40, after: 0)
        paraProps.keepNext = true
        paraProps.keepLines = true

        var runProps = RunProperties()
        runProps.fontSize = 26      // 13pt
        runProps.fontName = "Calibri Light"
        runProps.color = "2F5496"
        runProps.bold = true

        return Style(
            id: "Heading2",
            name: "heading 2",
            type: .paragraph,
            basedOn: "Normal",
            nextStyle: "Normal",
            paragraphProperties: paraProps,
            runProperties: runProps
        )
    }

    /// 標題 3
    static var heading3Style: Style {
        var paraProps = ParagraphProperties()
        paraProps.spacing = Spacing(before: 40, after: 0)
        paraProps.keepNext = true
        paraProps.keepLines = true

        var runProps = RunProperties()
        runProps.fontSize = 24      // 12pt
        runProps.fontName = "Calibri Light"
        runProps.color = "1F3763"
        runProps.bold = true

        return Style(
            id: "Heading3",
            name: "heading 3",
            type: .paragraph,
            basedOn: "Normal",
            nextStyle: "Normal",
            paragraphProperties: paraProps,
            runProperties: runProps
        )
    }

    /// 標題
    static var titleStyle: Style {
        var paraProps = ParagraphProperties()
        paraProps.spacing = Spacing(after: 0)
        paraProps.alignment = .center

        var runProps = RunProperties()
        runProps.fontSize = 56      // 28pt
        runProps.fontName = "Calibri Light"

        return Style(
            id: "Title",
            name: "Title",
            type: .paragraph,
            basedOn: "Normal",
            nextStyle: "Normal",
            paragraphProperties: paraProps,
            runProperties: runProps
        )
    }

    /// 副標題
    static var subtitleStyle: Style {
        var paraProps = ParagraphProperties()
        paraProps.spacing = Spacing(after: 160)
        paraProps.alignment = .center

        var runProps = RunProperties()
        runProps.fontSize = 24      // 12pt
        runProps.fontName = "Calibri"
        runProps.color = "5A5A5A"   // 灰色
        runProps.italic = true

        return Style(
            id: "Subtitle",
            name: "Subtitle",
            type: .paragraph,
            basedOn: "Normal",
            nextStyle: "Normal",
            paragraphProperties: paraProps,
            runProperties: runProps
        )
    }
}

// MARK: - XML 生成

extension Style {
    /// 轉換為 OOXML XML 字串
    func toXML() -> String {
        var parts: [String] = []

        // Style 開始標籤
        var attrs = "w:type=\"\(type.rawValue)\" w:styleId=\"\(id)\""
        if isDefault {
            attrs += " w:default=\"1\""
        }
        parts.append("<w:style \(attrs)>")

        // 名稱
        parts.append("<w:name w:val=\"\(name)\"/>")

        // 基於
        if let basedOn = basedOn {
            parts.append("<w:basedOn w:val=\"\(basedOn)\"/>")
        }

        // 下一段樣式
        if let nextStyle = nextStyle {
            parts.append("<w:next w:val=\"\(nextStyle)\"/>")
        }

        // 快速樣式
        if isQuickStyle {
            parts.append("<w:qFormat/>")
        }

        // 段落屬性
        if let paraProps = paragraphProperties {
            let propsXML = paraProps.toXML()
            if !propsXML.isEmpty {
                parts.append("<w:pPr>\(propsXML)</w:pPr>")
            }
        }

        // Run 屬性
        if let runProps = runProperties {
            let propsXML = runProps.toXML()
            if !propsXML.isEmpty {
                parts.append("<w:rPr>\(propsXML)</w:rPr>")
            }
        }

        parts.append("</w:style>")

        return parts.joined()
    }
}

// MARK: - Styles Collection XML

extension Array where Element == Style {
    /// 轉換為完整的 styles.xml 內容
    func toStylesXML() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        """

        // 預設樣式設定
        xml += """
        <w:docDefaults>
            <w:rPrDefault>
                <w:rPr>
                    <w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/>
                    <w:sz w:val="22"/>
                    <w:szCs w:val="22"/>
                </w:rPr>
            </w:rPrDefault>
            <w:pPrDefault>
                <w:pPr>
                    <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
                </w:pPr>
            </w:pPrDefault>
        </w:docDefaults>
        """

        // 所有樣式
        for style in self {
            xml += style.toXML()
        }

        xml += "</w:styles>"

        return xml
    }
}
