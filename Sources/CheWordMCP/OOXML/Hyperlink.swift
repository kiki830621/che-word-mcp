import Foundation

// MARK: - Hyperlink

/// 超連結（外部 URL 或內部書籤連結）
struct Hyperlink {
    var id: String              // 唯一識別碼（用於管理）
    var relationshipId: String? // 關係 ID（外部連結使用 rId）
    var anchor: String?         // 書籤名稱（內部連結使用）
    var url: String?            // 外部 URL
    var text: String            // 顯示文字
    var tooltip: String?        // 滑鼠懸停提示

    // 連結類型
    var type: HyperlinkType {
        if relationshipId != nil {
            return .external
        } else if anchor != nil {
            return .internal
        }
        return .external
    }

    init(id: String, text: String, url: String, relationshipId: String, tooltip: String? = nil) {
        self.id = id
        self.text = text
        self.url = url
        self.relationshipId = relationshipId
        self.anchor = nil
        self.tooltip = tooltip
    }

    init(id: String, text: String, anchor: String, tooltip: String? = nil) {
        self.id = id
        self.text = text
        self.anchor = anchor
        self.relationshipId = nil
        self.url = nil
        self.tooltip = tooltip
    }

    /// 建立外部連結
    static func external(id: String, text: String, url: String, relationshipId: String, tooltip: String? = nil) -> Hyperlink {
        return Hyperlink(id: id, text: text, url: url, relationshipId: relationshipId, tooltip: tooltip)
    }

    /// 建立內部連結（連到書籤）
    static func `internal`(id: String, text: String, bookmarkName: String, tooltip: String? = nil) -> Hyperlink {
        return Hyperlink(id: id, text: text, anchor: bookmarkName, tooltip: tooltip)
    }

    /// Relationship 類型（用於 .rels 檔案）
    static let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
}

/// 超連結類型
enum HyperlinkType {
    case external   // 外部 URL
    case `internal` // 內部書籤連結
}

// MARK: - Hyperlink Reference (用於 document.xml.rels)

/// 超連結關係（儲存在 document.xml.rels）
struct HyperlinkReference {
    var relationshipId: String  // rId
    var url: String            // 目標 URL

    init(relationshipId: String, url: String) {
        self.relationshipId = relationshipId
        self.url = url
    }
}

// MARK: - XML Generation

extension Hyperlink {
    /// 轉換為 OOXML XML（放在段落內）
    func toXML() -> String {
        var xml = "<w:hyperlink"

        // 外部連結使用 r:id，內部連結使用 w:anchor
        if let rId = relationshipId {
            xml += " r:id=\"\(rId)\""
        } else if let anchor = anchor {
            xml += " w:anchor=\"\(escapeXML(anchor))\""
        }

        // 提示文字
        if let tooltip = tooltip {
            xml += " w:tooltip=\"\(escapeXML(tooltip))\""
        }

        xml += ">"

        // 連結文字（帶有藍色底線樣式）
        xml += """
        <w:r>
            <w:rPr>
                <w:rStyle w:val="Hyperlink"/>
                <w:color w:val="0563C1"/>
                <w:u w:val="single"/>
            </w:rPr>
            <w:t xml:space="preserve">\(escapeXML(text))</w:t>
        </w:r>
        """

        xml += "</w:hyperlink>"
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
