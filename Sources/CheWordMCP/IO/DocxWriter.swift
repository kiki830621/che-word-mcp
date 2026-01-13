import Foundation

/// DOCX 檔案寫入器
struct DocxWriter {

    /// 將 WordDocument 寫入 .docx 檔案
    static func write(_ document: WordDocument, to url: URL) throws {
        // 1. 建立臨時目錄
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("che-word-mcp")
            .appendingPathComponent(UUID().uuidString)

        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)

        defer {
            ZipHelper.cleanup(tempDir)
        }

        // 2. 建立目錄結構
        try createDirectoryStructure(at: tempDir)

        // 3. 計算各種標記
        let hasNumbering = !document.numbering.abstractNums.isEmpty
        let hasHeaders = !document.headers.isEmpty
        let hasFooters = !document.footers.isEmpty

        // 4. 寫入各個 XML 檔案
        try writeContentTypes(to: tempDir, document: document)
        try writeRelationships(to: tempDir)
        try writeDocumentRelationships(to: tempDir, document: document)
        try writeDocument(document, to: tempDir)
        try writeStyles(document.styles, to: tempDir)
        try writeSettings(to: tempDir)
        try writeFontTable(to: tempDir)
        try writeCoreProperties(document.properties, to: tempDir)
        try writeAppProperties(to: tempDir)

        // 寫入編號定義（如果有）
        if hasNumbering {
            try writeNumbering(document.numbering, to: tempDir)
        }

        // 寫入頁首（如果有）
        if hasHeaders {
            for header in document.headers {
                try writeHeader(header, to: tempDir)
            }
        }

        // 寫入頁尾（如果有）
        if hasFooters {
            for footer in document.footers {
                try writeFooter(footer, to: tempDir)
            }
        }

        // 寫入圖片（如果有）
        if !document.images.isEmpty {
            try writeImages(document.images, to: tempDir)
        }

        // 寫入註解（如果有）
        if !document.comments.comments.isEmpty {
            try writeComments(document.comments, to: tempDir)
        }

        // 寫入腳註（如果有）
        if !document.footnotes.footnotes.isEmpty {
            try writeFootnotes(document.footnotes, to: tempDir)
        }

        // 寫入尾註（如果有）
        if !document.endnotes.endnotes.isEmpty {
            try writeEndnotes(document.endnotes, to: tempDir)
        }

        // 5. 壓縮成 ZIP
        try ZipHelper.zip(tempDir, to: url)
    }

    // MARK: - Directory Structure

    private static func createDirectoryStructure(at baseURL: URL) throws {
        let directories = [
            "_rels",
            "word",
            "word/_rels",
            "word/media",  // 圖片媒體目錄
            "docProps"
        ]

        for dir in directories {
            let dirURL = baseURL.appendingPathComponent(dir)
            try FileManager.default.createDirectory(at: dirURL, withIntermediateDirectories: true)
        }
    }

    // MARK: - Content Types

    private static func writeContentTypes(to baseURL: URL, document: WordDocument) throws {
        let hasNumbering = !document.numbering.abstractNums.isEmpty

        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Default Extension="png" ContentType="image/png"/>
            <Default Extension="jpeg" ContentType="image/jpeg"/>
            <Default Extension="jpg" ContentType="image/jpeg"/>
            <Default Extension="gif" ContentType="image/gif"/>
            <Default Extension="bmp" ContentType="image/bmp"/>
            <Default Extension="tiff" ContentType="image/tiff"/>
            <Default Extension="webp" ContentType="image/webp"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
            <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
            <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
            <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
            <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        """

        if hasNumbering {
            xml += """
                <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
            """
        }

        // 頁首
        for header in document.headers {
            xml += """
                <Override PartName="/word/\(header.fileName)" ContentType="\(Header.contentType)"/>
            """
        }

        // 頁尾
        for footer in document.footers {
            xml += """
                <Override PartName="/word/\(footer.fileName)" ContentType="\(Footer.contentType)"/>
            """
        }

        // 註解
        if !document.comments.comments.isEmpty {
            xml += """
                <Override PartName="/word/comments.xml" ContentType="\(CommentsCollection.contentType)"/>
            """
        }

        // 腳註
        if !document.footnotes.footnotes.isEmpty {
            xml += """
                <Override PartName="/word/footnotes.xml" ContentType="\(FootnotesCollection.contentType)"/>
            """
        }

        // 尾註
        if !document.endnotes.endnotes.isEmpty {
            xml += """
                <Override PartName="/word/endnotes.xml" ContentType="\(EndnotesCollection.contentType)"/>
            """
        }

        xml += "</Types>"

        let url = baseURL.appendingPathComponent("[Content_Types].xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Relationships

    private static func writeRelationships(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
        """

        let url = baseURL.appendingPathComponent("_rels/.rels")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    private static func writeDocumentRelationships(to baseURL: URL, document: WordDocument) throws {
        let hasNumbering = !document.numbering.abstractNums.isEmpty

        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
        """

        if hasNumbering {
            xml += """
                <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
            """
        }

        // 頁首關係
        for header in document.headers {
            xml += """
                <Relationship Id="\(header.id)" Type="\(Header.relationshipType)" Target="\(header.fileName)"/>
            """
        }

        // 頁尾關係
        for footer in document.footers {
            xml += """
                <Relationship Id="\(footer.id)" Type="\(Footer.relationshipType)" Target="\(footer.fileName)"/>
            """
        }

        // 圖片關係
        for image in document.images {
            xml += """
                <Relationship Id="\(image.id)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/\(image.fileName)"/>
            """
        }

        // 超連結關係（外部連結）
        for hyperlinkRef in document.hyperlinkReferences {
            xml += """
                <Relationship Id="\(hyperlinkRef.relationshipId)" Type="\(Hyperlink.relationshipType)" Target="\(escapeXML(hyperlinkRef.url))" TargetMode="External"/>
            """
        }

        // 註解關係
        if !document.comments.comments.isEmpty {
            // 計算下一個可用的 rId
            let baseId = document.numbering.abstractNums.isEmpty ? 4 : 5
            let usedCount = document.headers.count + document.footers.count + document.images.count + document.hyperlinkReferences.count
            let commentsRId = "rId\(baseId + usedCount)"
            xml += """
                <Relationship Id="\(commentsRId)" Type="\(CommentsCollection.relationshipType)" Target="comments.xml"/>
            """
        }

        // 腳註關係
        if !document.footnotes.footnotes.isEmpty {
            let baseId = document.numbering.abstractNums.isEmpty ? 4 : 5
            var usedCount = document.headers.count + document.footers.count + document.images.count + document.hyperlinkReferences.count
            if !document.comments.comments.isEmpty { usedCount += 1 }
            let footnotesRId = "rId\(baseId + usedCount)"
            xml += """
                <Relationship Id="\(footnotesRId)" Type="\(FootnotesCollection.relationshipType)" Target="footnotes.xml"/>
            """
        }

        // 尾註關係
        if !document.endnotes.endnotes.isEmpty {
            let baseId = document.numbering.abstractNums.isEmpty ? 4 : 5
            var usedCount = document.headers.count + document.footers.count + document.images.count + document.hyperlinkReferences.count
            if !document.comments.comments.isEmpty { usedCount += 1 }
            if !document.footnotes.footnotes.isEmpty { usedCount += 1 }
            let endnotesRId = "rId\(baseId + usedCount)"
            xml += """
                <Relationship Id="\(endnotesRId)" Type="\(EndnotesCollection.relationshipType)" Target="endnotes.xml"/>
            """
        }

        xml += "</Relationships>"

        let url = baseURL.appendingPathComponent("word/_rels/document.xml.rels")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Document

    private static func writeDocument(_ document: WordDocument, to baseURL: URL) throws {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        """

        // 段落和表格
        for child in document.body.children {
            switch child {
            case .paragraph(let para):
                xml += para.toXML()
            case .table(let table):
                xml += table.toXML()
            }
        }

        // 分節屬性（頁面設定）- 使用文件的 sectionProperties
        xml += document.sectionProperties.toXML()

        xml += "</w:body></w:document>"

        let url = baseURL.appendingPathComponent("word/document.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Styles

    private static func writeStyles(_ styles: [Style], to baseURL: URL) throws {
        let xml = styles.toStylesXML()
        let url = baseURL.appendingPathComponent("word/styles.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Numbering

    private static func writeNumbering(_ numbering: Numbering, to baseURL: URL) throws {
        let xml = numbering.toXML()
        let url = baseURL.appendingPathComponent("word/numbering.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Header

    private static func writeHeader(_ header: Header, to baseURL: URL) throws {
        let xml = header.toXML()
        let url = baseURL.appendingPathComponent("word/\(header.fileName)")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Footer

    private static func writeFooter(_ footer: Footer, to baseURL: URL) throws {
        let xml: String

        // 如果有指定頁碼格式，使用頁碼格式生成 XML
        if let format = footer.pageNumberFormat {
            xml = footer.toXMLWithPageNumber(format: format, alignment: footer.pageNumberAlignment)
        } else if footer.paragraphs.isEmpty {
            // 沒有段落也沒有頁碼格式，使用預設簡單頁碼
            xml = footer.toXMLWithPageNumber(format: .simple)
        } else {
            // 有段落內容，使用一般 XML 輸出
            xml = footer.toXML()
        }

        let url = baseURL.appendingPathComponent("word/\(footer.fileName)")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Images

    private static func writeImages(_ images: [ImageReference], to baseURL: URL) throws {
        for image in images {
            let url = baseURL.appendingPathComponent("word/media/\(image.fileName)")
            try image.data.write(to: url)
        }
    }

    // MARK: - Settings

    private static func writeSettings(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:defaultTabStop w:val="720"/>
            <w:characterSpacingControl w:val="doNotCompress"/>
            <w:compat>
                <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
            </w:compat>
        </w:settings>
        """

        let url = baseURL.appendingPathComponent("word/settings.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Font Table

    private static func writeFontTable(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:font w:name="Calibri">
                <w:panose1 w:val="020F0502020204030204"/>
                <w:charset w:val="00"/>
                <w:family w:val="swiss"/>
                <w:pitch w:val="variable"/>
            </w:font>
            <w:font w:name="Times New Roman">
                <w:panose1 w:val="02020603050405020304"/>
                <w:charset w:val="00"/>
                <w:family w:val="roman"/>
                <w:pitch w:val="variable"/>
            </w:font>
            <w:font w:name="Calibri Light">
                <w:panose1 w:val="020F0302020204030204"/>
                <w:charset w:val="00"/>
                <w:family w:val="swiss"/>
                <w:pitch w:val="variable"/>
            </w:font>
        </w:fonts>
        """

        let url = baseURL.appendingPathComponent("word/fontTable.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Core Properties

    private static func writeCoreProperties(_ props: DocumentProperties, to baseURL: URL) throws {
        let dateFormatter = ISO8601DateFormatter()

        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                           xmlns:dc="http://purl.org/dc/elements/1.1/"
                           xmlns:dcterms="http://purl.org/dc/terms/"
                           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        """

        if let title = props.title {
            xml += "<dc:title>\(escapeXML(title))</dc:title>"
        }
        if let subject = props.subject {
            xml += "<dc:subject>\(escapeXML(subject))</dc:subject>"
        }
        if let creator = props.creator {
            xml += "<dc:creator>\(escapeXML(creator))</dc:creator>"
        } else {
            xml += "<dc:creator>che-word-mcp</dc:creator>"
        }
        if let keywords = props.keywords {
            xml += "<cp:keywords>\(escapeXML(keywords))</cp:keywords>"
        }
        if let description = props.description {
            xml += "<dc:description>\(escapeXML(description))</dc:description>"
        }
        if let lastModifiedBy = props.lastModifiedBy {
            xml += "<cp:lastModifiedBy>\(escapeXML(lastModifiedBy))</cp:lastModifiedBy>"
        }
        if let revision = props.revision {
            xml += "<cp:revision>\(revision)</cp:revision>"
        } else {
            xml += "<cp:revision>1</cp:revision>"
        }

        let created = props.created ?? Date()
        xml += "<dcterms:created xsi:type=\"dcterms:W3CDTF\">\(dateFormatter.string(from: created))</dcterms:created>"

        let modified = props.modified ?? Date()
        xml += "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">\(dateFormatter.string(from: modified))</dcterms:modified>"

        xml += "</cp:coreProperties>"

        let url = baseURL.appendingPathComponent("docProps/core.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - App Properties

    private static func writeAppProperties(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <Application>che-word-mcp</Application>
            <AppVersion>1.0.0</AppVersion>
        </Properties>
        """

        let url = baseURL.appendingPathComponent("docProps/app.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Comments

    private static func writeComments(_ comments: CommentsCollection, to baseURL: URL) throws {
        let xml = comments.toXML()
        let url = baseURL.appendingPathComponent("word/comments.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Footnotes

    private static func writeFootnotes(_ footnotes: FootnotesCollection, to baseURL: URL) throws {
        let xml = footnotes.toXML()
        let url = baseURL.appendingPathComponent("word/footnotes.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Endnotes

    private static func writeEndnotes(_ endnotes: EndnotesCollection, to baseURL: URL) throws {
        let xml = endnotes.toXML()
        let url = baseURL.appendingPathComponent("word/endnotes.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Helpers

    private static func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}
