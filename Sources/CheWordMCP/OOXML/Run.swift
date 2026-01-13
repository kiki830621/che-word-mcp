import Foundation

/// 文字運行 (Run) - Word 文件中的最小文字單元
/// 一個 Run 包含具有相同格式的連續文字
struct Run {
    var text: String
    var properties: RunProperties
    var drawing: Drawing?  // 圖片繪圖元素

    init(text: String, properties: RunProperties = RunProperties()) {
        self.text = text
        self.properties = properties
        self.drawing = nil
    }
}

// MARK: - Run Properties

/// Run 格式屬性
struct RunProperties {
    var bold: Bool = false
    var italic: Bool = false
    var underline: UnderlineType?
    var strikethrough: Bool = false
    var fontSize: Int?              // 半點 (24 = 12pt)
    var fontName: String?
    var color: String?              // RGB hex (e.g., "FF0000")
    var highlight: HighlightColor?
    var verticalAlign: VerticalAlign?
    var characterSpacing: CharacterSpacing?  // 字元間距
    var textEffect: TextEffect?              // 文字效果
    var rawXML: String?                      // 原始 XML（用於進階功能如 SDT）

    init() {}

    init(bold: Bool = false,
         italic: Bool = false,
         underline: UnderlineType? = nil,
         strikethrough: Bool = false,
         fontSize: Int? = nil,
         fontName: String? = nil,
         color: String? = nil,
         highlight: HighlightColor? = nil,
         verticalAlign: VerticalAlign? = nil) {
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strikethrough = strikethrough
        self.fontSize = fontSize
        self.fontName = fontName
        self.color = color
        self.highlight = highlight
        self.verticalAlign = verticalAlign
    }

    /// 合併格式（覆蓋非 nil 值）
    mutating func merge(with other: RunProperties) {
        if other.bold { self.bold = true }
        if other.italic { self.italic = true }
        if let underline = other.underline { self.underline = underline }
        if other.strikethrough { self.strikethrough = true }
        if let fontSize = other.fontSize { self.fontSize = fontSize }
        if let fontName = other.fontName { self.fontName = fontName }
        if let color = other.color { self.color = color }
        if let highlight = other.highlight { self.highlight = highlight }
        if let verticalAlign = other.verticalAlign { self.verticalAlign = verticalAlign }
        if let characterSpacing = other.characterSpacing { self.characterSpacing = characterSpacing }
        if let textEffect = other.textEffect { self.textEffect = textEffect }
        if let rawXML = other.rawXML { self.rawXML = rawXML }
    }
}

// MARK: - Enums

/// 底線類型
enum UnderlineType: String, Codable {
    case single = "single"
    case double = "double"
    case dotted = "dotted"
    case dashed = "dash"
    case wave = "wave"
    case thick = "thick"
    case words = "words"        // 只在文字下，空格無底線
}

/// 螢光標記顏色
enum HighlightColor: String, Codable {
    case yellow = "yellow"
    case green = "green"
    case cyan = "cyan"
    case magenta = "magenta"
    case blue = "blue"
    case red = "red"
    case darkBlue = "darkBlue"
    case darkCyan = "darkCyan"
    case darkGreen = "darkGreen"
    case darkMagenta = "darkMagenta"
    case darkRed = "darkRed"
    case darkYellow = "darkYellow"
    case lightGray = "lightGray"
    case darkGray = "darkGray"
    case black = "black"
    case white = "white"
}

/// 垂直對齊（上標/下標）
enum VerticalAlign: String, Codable {
    case superscript = "superscript"
    case `subscript` = "subscript"
    case baseline = "baseline"
}

// MARK: - XML 生成

extension Run {
    /// 轉換為 OOXML XML 字串
    func toXML() -> String {
        // 如果有原始 XML，直接輸出（用於 SDT, TOC, 數學公式等）
        if let rawXML = properties.rawXML {
            return rawXML
        }

        var xml = "<w:r>"

        // Run Properties
        let propsXML = properties.toXML()
        if !propsXML.isEmpty {
            xml += "<w:rPr>\(propsXML)</w:rPr>"
        }

        // Drawing (圖片) - 如果有圖片，優先輸出圖片
        if let drawing = drawing {
            xml += drawing.toXML()
        } else {
            // Text (保留空格)
            xml += "<w:t xml:space=\"preserve\">\(escapeXML(text))</w:t>"
        }

        xml += "</w:r>"

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

extension RunProperties {
    /// 轉換為 OOXML XML 字串
    func toXML() -> String {
        var parts: [String] = []

        if bold {
            parts.append("<w:b/>")
        }
        if italic {
            parts.append("<w:i/>")
        }
        if let underline = underline {
            parts.append("<w:u w:val=\"\(underline.rawValue)\"/>")
        }
        if strikethrough {
            parts.append("<w:strike/>")
        }
        if let fontSize = fontSize {
            // OOXML 使用半點 (half-points)
            parts.append("<w:sz w:val=\"\(fontSize)\"/>")
            parts.append("<w:szCs w:val=\"\(fontSize)\"/>")  // 複雜文字大小
        }
        if let fontName = fontName {
            parts.append("<w:rFonts w:ascii=\"\(fontName)\" w:hAnsi=\"\(fontName)\" w:eastAsia=\"\(fontName)\" w:cs=\"\(fontName)\"/>")
        }
        if let color = color {
            parts.append("<w:color w:val=\"\(color)\"/>")
        }
        if let highlight = highlight {
            parts.append("<w:highlight w:val=\"\(highlight.rawValue)\"/>")
        }
        if let verticalAlign = verticalAlign {
            parts.append("<w:vertAlign w:val=\"\(verticalAlign.rawValue)\"/>")
        }
        if let characterSpacing = characterSpacing {
            parts.append(characterSpacing.toXML())
        }
        if let textEffect = textEffect {
            parts.append(textEffect.toXML())
        }

        return parts.joined()
    }
}
