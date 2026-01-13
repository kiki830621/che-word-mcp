import Foundation

// MARK: - Footer

/// 頁尾
struct Footer {
    var id: String           // 關係 ID (如 "rId11")
    var paragraphs: [Paragraph]
    var type: HeaderFooterType

    init(id: String, paragraphs: [Paragraph] = [], type: HeaderFooterType = .default) {
        self.id = id
        self.paragraphs = paragraphs
        self.type = type
    }

    /// 建立含單一文字的頁尾
    static func withText(_ text: String, id: String, type: HeaderFooterType = .default) -> Footer {
        var para = Paragraph(text: text)
        para.properties.alignment = .center
        return Footer(id: id, paragraphs: [para], type: type)
    }

    /// 建立含頁碼的頁尾
    static func withPageNumber(id: String, alignment: ParagraphAlignment = .center, format: PageNumberFormat = .simple, type: HeaderFooterType = .default) -> Footer {
        var para = Paragraph()
        para.properties.alignment = alignment
        return Footer(id: id, paragraphs: [para], type: type)
    }
}

// MARK: - Page Number Format

/// 頁碼格式
enum PageNumberFormat {
    case simple              // "1"
    case pageOfTotal         // "Page 1 of 10"
    case withDash            // "- 1 -"
    case withText(String)    // "第 1 頁" (使用 # 作為頁碼佔位符，如 "第#頁")
}

// MARK: - XML Generation

extension Footer {
    /// 轉換為完整的 footer.xml 內容
    func toXML() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        """

        for para in paragraphs {
            xml += para.toXML()
        }

        // 如果沒有段落，加一個空段落
        if paragraphs.isEmpty {
            xml += "<w:p/>"
        }

        xml += "</w:ftr>"
        return xml
    }

    /// 轉換為含頁碼的頁尾 XML
    func toXMLWithPageNumber(format: PageNumberFormat = .simple, alignment: ParagraphAlignment = .center) -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:p>
        <w:pPr><w:jc w:val="\(alignment.rawValue)"/></w:pPr>
        """

        // 根據格式添加內容
        switch format {
        case .simple:
            xml += pageFieldXML()

        case .pageOfTotal:
            xml += "<w:r><w:t xml:space=\"preserve\">Page </w:t></w:r>"
            xml += pageFieldXML()
            xml += "<w:r><w:t xml:space=\"preserve\"> of </w:t></w:r>"
            xml += numPagesFieldXML()

        case .withDash:
            xml += "<w:r><w:t xml:space=\"preserve\">- </w:t></w:r>"
            xml += pageFieldXML()
            xml += "<w:r><w:t xml:space=\"preserve\"> -</w:t></w:r>"

        case .withText(let template):
            let parts = template.components(separatedBy: "#")
            if parts.count >= 1 && !parts[0].isEmpty {
                xml += "<w:r><w:t xml:space=\"preserve\">\(escapeXML(parts[0]))</w:t></w:r>"
            }
            xml += pageFieldXML()
            if parts.count >= 2 && !parts[1].isEmpty {
                xml += "<w:r><w:t xml:space=\"preserve\">\(escapeXML(parts[1]))</w:t></w:r>"
            }
        }

        xml += "</w:p></w:ftr>"
        return xml
    }

    /// PAGE 欄位 XML
    private func pageFieldXML() -> String {
        return """
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        """
    }

    /// NUMPAGES 欄位 XML
    private func numPagesFieldXML() -> String {
        return """
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        """
    }

    /// XML 跳脫
    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }

    /// 取得檔案名稱
    var fileName: String {
        switch type {
        case .default: return "footer1.xml"
        case .first: return "footerFirst.xml"
        case .even: return "footerEven.xml"
        }
    }

    /// 取得關係類型
    static var relationshipType: String {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    }

    /// 取得內容類型
    static var contentType: String {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    }
}
