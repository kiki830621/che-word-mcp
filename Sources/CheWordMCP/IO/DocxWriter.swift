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

        // 3. 寫入各個 XML 檔案
        try writeContentTypes(to: tempDir)
        try writeRelationships(to: tempDir)
        try writeDocumentRelationships(to: tempDir)
        try writeDocument(document, to: tempDir)
        try writeStyles(document.styles, to: tempDir)
        try writeSettings(to: tempDir)
        try writeFontTable(to: tempDir)
        try writeCoreProperties(document.properties, to: tempDir)
        try writeAppProperties(to: tempDir)

        // 4. 壓縮成 ZIP
        try ZipHelper.zip(tempDir, to: url)
    }

    // MARK: - Directory Structure

    private static func createDirectoryStructure(at baseURL: URL) throws {
        let directories = [
            "_rels",
            "word",
            "word/_rels",
            "docProps"
        ]

        for dir in directories {
            let dirURL = baseURL.appendingPathComponent(dir)
            try FileManager.default.createDirectory(at: dirURL, withIntermediateDirectories: true)
        }
    }

    // MARK: - Content Types

    private static func writeContentTypes(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
            <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
            <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
            <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
            <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        </Types>
        """

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

    private static func writeDocumentRelationships(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
        </Relationships>
        """

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

        // 分節屬性（頁面設定）
        xml += """
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
            <w:docGrid w:linePitch="360"/>
        </w:sectPr>
        """

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
