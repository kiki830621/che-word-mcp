import Foundation

/// 表格 (Table) - Word 文件中的表格結構
struct Table {
    var rows: [TableRow]
    var properties: TableProperties

    init(rows: [TableRow] = [], properties: TableProperties = TableProperties()) {
        self.rows = rows
        self.properties = properties
    }

    /// 便利初始化器：建立指定行列的空表格
    init(rowCount: Int, columnCount: Int, properties: TableProperties = TableProperties()) {
        self.properties = properties
        self.rows = (0..<rowCount).map { _ in
            TableRow(cells: (0..<columnCount).map { _ in TableCell() })
        }
    }

    /// 取得表格純文字
    func getText() -> String {
        return rows.map { row in
            row.cells.map { cell in
                cell.getText()
            }.joined(separator: "\t")
        }.joined(separator: "\n")
    }

    /// 轉換為 Markdown 表格
    func toMarkdown() -> String {
        guard !rows.isEmpty else { return "" }

        var result: [String] = []

        // 表頭（第一行）
        if let firstRow = rows.first {
            let headers = firstRow.cells.map { $0.getText() }
            result.append("| " + headers.joined(separator: " | ") + " |")
            result.append("|" + headers.map { _ in "---" }.joined(separator: "|") + "|")
        }

        // 資料行
        for row in rows.dropFirst() {
            let cells = row.cells.map { $0.getText() }
            result.append("| " + cells.joined(separator: " | ") + " |")
        }

        return result.joined(separator: "\n")
    }
}

// MARK: - Table Row

/// 表格行
struct TableRow {
    var cells: [TableCell]
    var properties: TableRowProperties

    init(cells: [TableCell] = [], properties: TableRowProperties = TableRowProperties()) {
        self.cells = cells
        self.properties = properties
    }
}

/// 表格行屬性
struct TableRowProperties {
    var height: Int?                // 行高 (twips)
    var heightRule: HeightRule?     // 行高規則
    var isHeader: Bool = false      // 是否為表頭行（每頁重複）
    var cantSplit: Bool = false     // 禁止跨頁分割

    init() {}
}

/// 行高規則
enum HeightRule: String, Codable {
    case auto = "auto"
    case exact = "exact"
    case atLeast = "atLeast"
}

// MARK: - Table Cell

/// 表格儲存格
struct TableCell {
    var paragraphs: [Paragraph]
    var properties: TableCellProperties

    init() {
        self.paragraphs = [Paragraph()]
        self.properties = TableCellProperties()
    }

    init(paragraphs: [Paragraph], properties: TableCellProperties = TableCellProperties()) {
        self.paragraphs = paragraphs.isEmpty ? [Paragraph()] : paragraphs
        self.properties = properties
    }

    /// 便利初始化器：用文字建立儲存格
    init(text: String) {
        self.paragraphs = [Paragraph(text: text)]
        self.properties = TableCellProperties()
    }

    /// 取得儲存格純文字
    func getText() -> String {
        return paragraphs.map { $0.getText() }.joined(separator: "\n")
    }
}

/// 表格儲存格屬性
struct TableCellProperties {
    var width: Int?                     // 寬度 (twips)
    var widthType: WidthType?           // 寬度類型
    var verticalAlignment: CellVerticalAlignment?
    var gridSpan: Int?                  // 水平合併（跨幾欄）
    var verticalMerge: VerticalMerge?   // 垂直合併
    var borders: CellBorders?           // 邊框
    var shading: CellShading?           // 底色

    init() {}
}

/// 寬度類型
enum WidthType: String, Codable {
    case auto = "auto"
    case dxa = "dxa"        // twips
    case pct = "pct"        // 百分比 (50 = 50%)
    case nil_ = "nil"       // 無寬度
}

/// 儲存格垂直對齊
enum CellVerticalAlignment: String, Codable {
    case top = "top"
    case center = "center"
    case bottom = "bottom"
}

/// 垂直合併
enum VerticalMerge: String, Codable {
    case restart = "restart"    // 合併的第一個儲存格
    case `continue` = "continue" // 被合併的儲存格
}

/// 儲存格邊框
struct CellBorders {
    var top: Border?
    var bottom: Border?
    var left: Border?
    var right: Border?

    init(top: Border? = nil, bottom: Border? = nil, left: Border? = nil, right: Border? = nil) {
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
    }

    /// 便利方法：建立四邊相同邊框
    static func all(_ border: Border) -> CellBorders {
        CellBorders(top: border, bottom: border, left: border, right: border)
    }
}

/// 邊框
struct Border {
    var style: BorderStyle
    var size: Int           // 1/8 點
    var color: String       // RGB hex

    init(style: BorderStyle = .single, size: Int = 4, color: String = "000000") {
        self.style = style
        self.size = size
        self.color = color
    }
}

/// 邊框樣式
enum BorderStyle: String, Codable {
    case single = "single"
    case double = "double"
    case dotted = "dotted"
    case dashed = "dashed"
    case thick = "thick"
    case nil_ = "nil"           // 無邊框
}

/// 儲存格底色
struct CellShading {
    var fill: String            // 背景色 RGB hex
    var color: String?          // 前景色（用於圖案）
    var pattern: ShadingPattern?

    init(fill: String, color: String? = nil, pattern: ShadingPattern? = nil) {
        self.fill = fill
        self.color = color
        self.pattern = pattern
    }

    /// 便利方法：純色背景
    static func solid(_ color: String) -> CellShading {
        CellShading(fill: color, pattern: .clear)
    }
}

/// 底色圖案
enum ShadingPattern: String, Codable {
    case clear = "clear"
    case solid = "solid"
    case horzStripe = "horzStripe"
    case vertStripe = "vertStripe"
    case diagStripe = "diagStripe"
}

// MARK: - Table Properties

/// 表格屬性
struct TableProperties {
    var width: Int?                     // 表格寬度
    var widthType: WidthType?
    var alignment: Alignment?           // 表格對齊
    var borders: TableBorders?          // 表格邊框
    var cellMargins: TableCellMargins?  // 預設儲存格邊距
    var layout: TableLayout?            // 版面配置

    init() {}
}

/// 表格邊框
struct TableBorders {
    var top: Border?
    var bottom: Border?
    var left: Border?
    var right: Border?
    var insideH: Border?    // 內部水平線
    var insideV: Border?    // 內部垂直線

    init() {}

    /// 便利方法：建立全邊框
    static func all(_ border: Border) -> TableBorders {
        var borders = TableBorders()
        borders.top = border
        borders.bottom = border
        borders.left = border
        borders.right = border
        borders.insideH = border
        borders.insideV = border
        return borders
    }
}

/// 表格儲存格邊距
struct TableCellMargins {
    var top: Int?
    var bottom: Int?
    var left: Int?
    var right: Int?

    init(top: Int? = nil, bottom: Int? = nil, left: Int? = nil, right: Int? = nil) {
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
    }

    /// 便利方法：四邊相同邊距
    static func all(_ margin: Int) -> TableCellMargins {
        TableCellMargins(top: margin, bottom: margin, left: margin, right: margin)
    }
}

/// 表格版面配置
enum TableLayout: String, Codable {
    case fixed = "fixed"        // 固定欄寬
    case autofit = "autofit"    // 自動調整
}

// MARK: - XML 生成

extension Table {
    /// 轉換為 OOXML XML 字串
    func toXML() -> String {
        var xml = "<w:tbl>"

        // Table Properties
        xml += properties.toXML()

        // Table Grid (欄位定義)
        xml += "<w:tblGrid>"
        if let firstRow = rows.first {
            for cell in firstRow.cells {
                let width = cell.properties.width ?? 2000
                xml += "<w:gridCol w:w=\"\(width)\"/>"
            }
        }
        xml += "</w:tblGrid>"

        // Rows
        for row in rows {
            xml += row.toXML()
        }

        xml += "</w:tbl>"
        return xml
    }
}

extension TableProperties {
    func toXML() -> String {
        var parts: [String] = ["<w:tblPr>"]

        // 寬度
        if let width = width {
            let type = widthType?.rawValue ?? "dxa"
            parts.append("<w:tblW w:w=\"\(width)\" w:type=\"\(type)\"/>")
        }

        // 對齊
        if let alignment = alignment {
            parts.append("<w:jc w:val=\"\(alignment.rawValue)\"/>")
        }

        // 版面配置
        if let layout = layout {
            parts.append("<w:tblLayout w:type=\"\(layout.rawValue)\"/>")
        }

        // 邊框
        if let borders = borders {
            parts.append(borders.toXML())
        }

        // 儲存格邊距
        if let margins = cellMargins {
            parts.append("<w:tblCellMar>")
            if let top = margins.top { parts.append("<w:top w:w=\"\(top)\" w:type=\"dxa\"/>") }
            if let bottom = margins.bottom { parts.append("<w:bottom w:w=\"\(bottom)\" w:type=\"dxa\"/>") }
            if let left = margins.left { parts.append("<w:left w:w=\"\(left)\" w:type=\"dxa\"/>") }
            if let right = margins.right { parts.append("<w:right w:w=\"\(right)\" w:type=\"dxa\"/>") }
            parts.append("</w:tblCellMar>")
        }

        parts.append("</w:tblPr>")
        return parts.joined()
    }
}

extension TableBorders {
    func toXML() -> String {
        var parts: [String] = ["<w:tblBorders>"]

        if let top = top { parts.append(top.toXML(name: "top")) }
        if let bottom = bottom { parts.append(bottom.toXML(name: "bottom")) }
        if let left = left { parts.append(left.toXML(name: "left")) }
        if let right = right { parts.append(right.toXML(name: "right")) }
        if let insideH = insideH { parts.append(insideH.toXML(name: "insideH")) }
        if let insideV = insideV { parts.append(insideV.toXML(name: "insideV")) }

        parts.append("</w:tblBorders>")
        return parts.joined()
    }
}

extension Border {
    func toXML(name: String) -> String {
        return "<w:\(name) w:val=\"\(style.rawValue)\" w:sz=\"\(size)\" w:color=\"\(color)\"/>"
    }
}

extension TableRow {
    func toXML() -> String {
        var xml = "<w:tr>"

        // Row Properties
        if properties.height != nil || properties.isHeader || properties.cantSplit {
            xml += "<w:trPr>"
            if let height = properties.height {
                let rule = properties.heightRule?.rawValue ?? "auto"
                xml += "<w:trHeight w:val=\"\(height)\" w:hRule=\"\(rule)\"/>"
            }
            if properties.isHeader {
                xml += "<w:tblHeader/>"
            }
            if properties.cantSplit {
                xml += "<w:cantSplit/>"
            }
            xml += "</w:trPr>"
        }

        // Cells
        for cell in cells {
            xml += cell.toXML()
        }

        xml += "</w:tr>"
        return xml
    }
}

extension TableCell {
    func toXML() -> String {
        var xml = "<w:tc>"

        // Cell Properties
        xml += properties.toXML()

        // Paragraphs (每個儲存格至少需要一個段落)
        if paragraphs.isEmpty {
            xml += Paragraph().toXML()
        } else {
            for para in paragraphs {
                xml += para.toXML()
            }
        }

        xml += "</w:tc>"
        return xml
    }
}

extension TableCellProperties {
    func toXML() -> String {
        var parts: [String] = ["<w:tcPr>"]

        // 寬度
        if let width = width {
            let type = widthType?.rawValue ?? "dxa"
            parts.append("<w:tcW w:w=\"\(width)\" w:type=\"\(type)\"/>")
        }

        // 水平合併
        if let gridSpan = gridSpan, gridSpan > 1 {
            parts.append("<w:gridSpan w:val=\"\(gridSpan)\"/>")
        }

        // 垂直合併
        if let vMerge = verticalMerge {
            parts.append("<w:vMerge w:val=\"\(vMerge.rawValue)\"/>")
        }

        // 垂直對齊
        if let vAlign = verticalAlignment {
            parts.append("<w:vAlign w:val=\"\(vAlign.rawValue)\"/>")
        }

        // 邊框
        if let borders = borders {
            parts.append("<w:tcBorders>")
            if let top = borders.top { parts.append(top.toXML(name: "top")) }
            if let bottom = borders.bottom { parts.append(bottom.toXML(name: "bottom")) }
            if let left = borders.left { parts.append(left.toXML(name: "left")) }
            if let right = borders.right { parts.append(right.toXML(name: "right")) }
            parts.append("</w:tcBorders>")
        }

        // 底色
        if let shading = shading {
            var attrs = "w:fill=\"\(shading.fill)\""
            if let color = shading.color { attrs += " w:color=\"\(color)\"" }
            if let pattern = shading.pattern { attrs += " w:val=\"\(pattern.rawValue)\"" }
            parts.append("<w:shd \(attrs)/>")
        }

        parts.append("</w:tcPr>")
        return parts.joined()
    }
}
