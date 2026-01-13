import Foundation

// MARK: - Header

/// 頁首
struct Header {
    var id: String           // 關係 ID (如 "rId10")
    var paragraphs: [Paragraph]
    var type: HeaderFooterType

    init(id: String, paragraphs: [Paragraph] = [], type: HeaderFooterType = .default) {
        self.id = id
        self.paragraphs = paragraphs
        self.type = type
    }

    /// 建立含單一文字的頁首
    static func withText(_ text: String, id: String, type: HeaderFooterType = .default) -> Header {
        var para = Paragraph(text: text)
        para.properties.alignment = .center
        return Header(id: id, paragraphs: [para], type: type)
    }

    /// 建立含頁碼的頁首
    static func withPageNumber(id: String, alignment: ParagraphAlignment = .center, type: HeaderFooterType = .default) -> Header {
        // 頁碼會在 XML 生成時處理
        var para = Paragraph()
        para.properties.alignment = alignment
        return Header(id: id, paragraphs: [para], type: type)
    }
}

// MARK: - Header/Footer Type

/// 頁首/頁尾類型
enum HeaderFooterType: String {
    case `default` = "default"  // 預設（奇數頁/所有頁）
    case first = "first"        // 首頁
    case even = "even"          // 偶數頁
}

// MARK: - XML Generation

extension Header {
    /// 轉換為完整的 header.xml 內容
    func toXML() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        """

        for para in paragraphs {
            xml += para.toXML()
        }

        // 如果沒有段落，加一個空段落
        if paragraphs.isEmpty {
            xml += "<w:p/>"
        }

        xml += "</w:hdr>"
        return xml
    }

    /// 轉換為含頁碼的頁首 XML
    func toXMLWithPageNumber() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:p>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        """

        xml += pageFieldXML()

        xml += "</w:p></w:hdr>"
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

    /// 取得檔案名稱
    var fileName: String {
        switch type {
        case .default: return "header1.xml"
        case .first: return "headerFirst.xml"
        case .even: return "headerEven.xml"
        }
    }

    /// 取得關係類型
    static var relationshipType: String {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    }

    /// 取得內容類型
    static var contentType: String {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    }
}
