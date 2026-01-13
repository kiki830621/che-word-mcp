import Foundation
import MCP

/// Word MCP Server - Swift OOXML Word 文件處理
actor WordMCPServer {
    private let server: Server
    private let transport: StdioTransport

    /// 目前開啟的文件 (doc_id -> WordDocument)
    private var openDocuments: [String: WordDocument] = [:]

    init() {
        self.server = Server(
            name: "che-word-mcp",
            version: "1.0.0"
        )
        self.transport = StdioTransport()
    }

    func run() async throws {
        // 註冊 Tool handlers
        await registerToolHandlers()

        // 啟動 server
        try await server.start(transport: transport)

        // 等待完成
        await server.waitUntilCompleted()
    }

    private func registerToolHandlers() async {
        let tools = allTools

        // 列出所有工具
        await server.withMethodHandler(ListTools.self) { [tools] _ in
            ListTools.Result(tools: tools)
        }

        // 處理工具呼叫
        await server.withMethodHandler(CallTool.self) { [weak self] params in
            guard let self = self else {
                return CallTool.Result(content: [.text("Server unavailable")], isError: true)
            }
            return try await self.handleToolCall(params)
        }
    }

    // MARK: - Tools Definition

    private var allTools: [Tool] {
        [
            // 文件管理
            Tool(
                name: "create_document",
                description: "建立新的 Word 文件 (.docx)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼，用於後續操作")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "open_document",
                description: "開啟現有的 Word 文件 (.docx)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("文件路徑")
                        ]),
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼，用於後續操作")
                        ])
                    ]),
                    "required": .array([.string("path"), .string("doc_id")])
                ])
            ),
            Tool(
                name: "save_document",
                description: "儲存 Word 文件 (.docx)",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("儲存路徑")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("path")])
                ])
            ),
            Tool(
                name: "close_document",
                description: "關閉已開啟的文件",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "list_open_documents",
                description: "列出所有已開啟的文件",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([:])
                ])
            ),
            Tool(
                name: "get_document_info",
                description: "取得文件資訊（段落數、字數等）",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),

            // 內容操作
            Tool(
                name: "get_text",
                description: "取得文件的純文字內容",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "get_paragraphs",
                description: "取得所有段落（含格式資訊）",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ])
                    ]),
                    "required": .array([.string("doc_id")])
                ])
            ),
            Tool(
                name: "insert_paragraph",
                description: "插入新段落",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("段落文字內容")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("插入位置（從 0 開始），不指定則加到最後")
                        ]),
                        "style": .object([
                            "type": .string("string"),
                            "description": .string("段落樣式（如 Heading1, Normal）")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("text")])
                ])
            ),
            Tool(
                name: "update_paragraph",
                description: "更新現有段落的內容",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("段落索引（從 0 開始）")
                        ]),
                        "text": .object([
                            "type": .string("string"),
                            "description": .string("新的段落文字")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("index"), .string("text")])
                ])
            ),
            Tool(
                name: "delete_paragraph",
                description: "刪除段落",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("段落索引（從 0 開始）")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("index")])
                ])
            ),
            Tool(
                name: "replace_text",
                description: "搜尋並取代文字",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "find": .object([
                            "type": .string("string"),
                            "description": .string("要搜尋的文字")
                        ]),
                        "replace": .object([
                            "type": .string("string"),
                            "description": .string("取代後的文字")
                        ]),
                        "all": .object([
                            "type": .string("boolean"),
                            "description": .string("是否取代所有符合項目（預設 true）")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("find"), .string("replace")])
                ])
            ),

            // 格式化
            Tool(
                name: "format_text",
                description: "格式化指定段落的文字（粗體、斜體、顏色等）",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("段落索引")
                        ]),
                        "bold": .object([
                            "type": .string("boolean"),
                            "description": .string("粗體")
                        ]),
                        "italic": .object([
                            "type": .string("boolean"),
                            "description": .string("斜體")
                        ]),
                        "underline": .object([
                            "type": .string("boolean"),
                            "description": .string("底線")
                        ]),
                        "font_size": .object([
                            "type": .string("integer"),
                            "description": .string("字型大小（點數，如 12）")
                        ]),
                        "font_name": .object([
                            "type": .string("string"),
                            "description": .string("字型名稱（如 Arial, Times New Roman）")
                        ]),
                        "color": .object([
                            "type": .string("string"),
                            "description": .string("文字顏色（RGB 十六進位，如 FF0000 表示紅色）")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),
            Tool(
                name: "set_paragraph_format",
                description: "設定段落格式（對齊、間距等）",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("段落索引")
                        ]),
                        "alignment": .object([
                            "type": .string("string"),
                            "description": .string("對齊方式：left, center, right, both")
                        ]),
                        "line_spacing": .object([
                            "type": .string("number"),
                            "description": .string("行距（倍數，如 1.5）")
                        ]),
                        "space_before": .object([
                            "type": .string("integer"),
                            "description": .string("段前間距（點數）")
                        ]),
                        "space_after": .object([
                            "type": .string("integer"),
                            "description": .string("段後間距（點數）")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index")])
                ])
            ),
            Tool(
                name: "apply_style",
                description: "套用內建樣式到段落",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "paragraph_index": .object([
                            "type": .string("integer"),
                            "description": .string("段落索引")
                        ]),
                        "style": .object([
                            "type": .string("string"),
                            "description": .string("樣式名稱（如 Heading1, Heading2, Normal, Title）")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("paragraph_index"), .string("style")])
                ])
            ),

            // 表格
            Tool(
                name: "insert_table",
                description: "插入表格",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "rows": .object([
                            "type": .string("integer"),
                            "description": .string("列數")
                        ]),
                        "cols": .object([
                            "type": .string("integer"),
                            "description": .string("欄數")
                        ]),
                        "data": .object([
                            "type": .string("array"),
                            "description": .string("表格資料（二維陣列）")
                        ]),
                        "index": .object([
                            "type": .string("integer"),
                            "description": .string("插入位置")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("rows"), .string("cols")])
                ])
            ),

            // 匯出
            Tool(
                name: "export_text",
                description: "匯出文件為純文字",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("匯出路徑")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("path")])
                ])
            ),
            Tool(
                name: "export_markdown",
                description: "匯出文件為 Markdown 格式",
                inputSchema: .object([
                    "type": .string("object"),
                    "properties": .object([
                        "doc_id": .object([
                            "type": .string("string"),
                            "description": .string("文件識別碼")
                        ]),
                        "path": .object([
                            "type": .string("string"),
                            "description": .string("匯出路徑")
                        ])
                    ]),
                    "required": .array([.string("doc_id"), .string("path")])
                ])
            )
        ]
    }

    // MARK: - Tool Handler

    private func handleToolCall(_ params: CallTool.Parameters) async throws -> CallTool.Result {
        let name = params.name
        let args = params.arguments ?? [:]

        do {
            let result = try await executeToolTask(name: name, args: args)
            return CallTool.Result(content: [.text(result)])
        } catch {
            return CallTool.Result(
                content: [.text("Error: \(error.localizedDescription)")],
                isError: true
            )
        }
    }

    private func executeToolTask(name: String, args: [String: Value]) async throws -> String {
        switch name {
        // 文件管理
        case "create_document":
            return try await createDocument(args: args)
        case "open_document":
            return try await openDocument(args: args)
        case "save_document":
            return try await saveDocument(args: args)
        case "close_document":
            return try await closeDocument(args: args)
        case "list_open_documents":
            return await listOpenDocuments()
        case "get_document_info":
            return try await getDocumentInfo(args: args)

        // 內容操作
        case "get_text":
            return try await getText(args: args)
        case "get_paragraphs":
            return try await getParagraphs(args: args)
        case "insert_paragraph":
            return try await insertParagraph(args: args)
        case "update_paragraph":
            return try await updateParagraph(args: args)
        case "delete_paragraph":
            return try await deleteParagraph(args: args)
        case "replace_text":
            return try await replaceText(args: args)

        // 格式化
        case "format_text":
            return try await formatText(args: args)
        case "set_paragraph_format":
            return try await setParagraphFormat(args: args)
        case "apply_style":
            return try await applyStyle(args: args)

        // 表格
        case "insert_table":
            return try await insertTable(args: args)

        // 匯出
        case "export_text":
            return try await exportText(args: args)
        case "export_markdown":
            return try await exportMarkdown(args: args)

        default:
            throw WordError.unknownTool(name)
        }
    }

    // MARK: - Document Management

    private func createDocument(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }

        let doc = WordDocument()
        openDocuments[docId] = doc

        return "Created new document with id: \(docId)"
    }

    private func openDocument(args: [String: Value]) async throws -> String {
        guard let path = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }

        let url = URL(fileURLWithPath: path)
        let doc = try DocxReader.read(from: url)
        openDocuments[docId] = doc

        return "Opened document '\(path)' with id: \(docId)"
    }

    private func saveDocument(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let path = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let url = URL(fileURLWithPath: path)
        try DocxWriter.write(doc, to: url)

        return "Saved document to: \(path)"
    }

    private func closeDocument(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }

        openDocuments.removeValue(forKey: docId)
        return "Closed document: \(docId)"
    }

    private func listOpenDocuments() async -> String {
        if openDocuments.isEmpty {
            return "No documents currently open"
        }

        let ids = openDocuments.keys.sorted()
        return "Open documents:\n" + ids.map { "- \($0)" }.joined(separator: "\n")
    }

    private func getDocumentInfo(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let info = doc.getInfo()
        return """
        Document Info (\(docId)):
        - Paragraphs: \(info.paragraphCount)
        - Characters: \(info.characterCount)
        - Words: \(info.wordCount)
        - Tables: \(info.tableCount)
        """
    }

    // MARK: - Content Operations

    private func getText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        return doc.getText()
    }

    private func getParagraphs(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let paragraphs = doc.getParagraphs()
        if paragraphs.isEmpty {
            return "No paragraphs in document"
        }

        var result = "Paragraphs:\n"
        for (index, para) in paragraphs.enumerated() {
            let style = para.properties.style ?? "Normal"
            let preview = String(para.getText().prefix(50))
            result += "[\(index)] (\(style)) \(preview)...\n"
        }
        return result
    }

    private func insertParagraph(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let index = args["index"]?.intValue
        let style = args["style"]?.stringValue

        var para = Paragraph(text: text)
        if let style = style {
            para.properties.style = style
        }

        if let index = index {
            doc.insertParagraph(para, at: index)
        } else {
            doc.appendParagraph(para)
        }

        openDocuments[docId] = doc

        return "Inserted paragraph at index \(index ?? doc.getParagraphs().count - 1)"
    }

    private func updateParagraph(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let index = args["index"]?.intValue else {
            throw WordError.missingParameter("index")
        }
        guard let text = args["text"]?.stringValue else {
            throw WordError.missingParameter("text")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        try doc.updateParagraph(at: index, text: text)
        openDocuments[docId] = doc

        return "Updated paragraph at index \(index)"
    }

    private func deleteParagraph(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let index = args["index"]?.intValue else {
            throw WordError.missingParameter("index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        try doc.deleteParagraph(at: index)
        openDocuments[docId] = doc

        return "Deleted paragraph at index \(index)"
    }

    private func replaceText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let find = args["find"]?.stringValue else {
            throw WordError.missingParameter("find")
        }
        guard let replace = args["replace"]?.stringValue else {
            throw WordError.missingParameter("replace")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let replaceAll = args["all"]?.boolValue ?? true
        let count = doc.replaceText(find: find, with: replace, all: replaceAll)
        openDocuments[docId] = doc

        return "Replaced \(count) occurrence(s) of '\(find)' with '\(replace)'"
    }

    // MARK: - Formatting

    private func formatText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        var format = RunProperties()
        if let bold = args["bold"]?.boolValue { format.bold = bold }
        if let italic = args["italic"]?.boolValue { format.italic = italic }
        if let underline = args["underline"]?.boolValue { format.underline = underline ? .single : nil }
        if let fontSize = args["font_size"]?.intValue { format.fontSize = fontSize * 2 } // 轉換為半點
        if let fontName = args["font_name"]?.stringValue { format.fontName = fontName }
        if let color = args["color"]?.stringValue { format.color = color }

        try doc.formatParagraph(at: paragraphIndex, with: format)
        openDocuments[docId] = doc

        return "Applied formatting to paragraph \(paragraphIndex)"
    }

    private func setParagraphFormat(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        var props = ParagraphProperties()
        if let alignment = args["alignment"]?.stringValue {
            props.alignment = Alignment(rawValue: alignment)
        }
        if let lineSpacing = args["line_spacing"]?.doubleValue {
            props.spacing = Spacing(line: Int(lineSpacing * 240)) // 轉換為 1/240 點
        }
        if let spaceBefore = args["space_before"]?.intValue {
            if props.spacing == nil { props.spacing = Spacing() }
            props.spacing?.before = spaceBefore * 20 // 轉換為 1/20 點
        }
        if let spaceAfter = args["space_after"]?.intValue {
            if props.spacing == nil { props.spacing = Spacing() }
            props.spacing?.after = spaceAfter * 20
        }

        try doc.setParagraphFormat(at: paragraphIndex, properties: props)
        openDocuments[docId] = doc

        return "Applied paragraph format to index \(paragraphIndex)"
    }

    private func applyStyle(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let paragraphIndex = args["paragraph_index"]?.intValue else {
            throw WordError.missingParameter("paragraph_index")
        }
        guard let style = args["style"]?.stringValue else {
            throw WordError.missingParameter("style")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        try doc.applyStyle(at: paragraphIndex, style: style)
        openDocuments[docId] = doc

        return "Applied style '\(style)' to paragraph \(paragraphIndex)"
    }

    // MARK: - Table

    private func insertTable(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let rows = args["rows"]?.intValue else {
            throw WordError.missingParameter("rows")
        }
        guard let cols = args["cols"]?.intValue else {
            throw WordError.missingParameter("cols")
        }
        guard var doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        var table = Table(rowCount: rows, columnCount: cols)

        // 如果有提供資料，填入表格
        if let dataArray = args["data"]?.arrayValue {
            for (rowIndex, rowData) in dataArray.enumerated() {
                if let rowArray = rowData.arrayValue {
                    for (colIndex, cellData) in rowArray.enumerated() {
                        if let text = cellData.stringValue,
                           rowIndex < table.rows.count && colIndex < table.rows[rowIndex].cells.count {
                            table.rows[rowIndex].cells[colIndex] = TableCell(text: text)
                        }
                    }
                }
            }
        }

        let index = args["index"]?.intValue
        if let index = index {
            doc.insertTable(table, at: index)
        } else {
            doc.appendTable(table)
        }

        openDocuments[docId] = doc

        return "Inserted \(rows)x\(cols) table"
    }

    // MARK: - Export

    private func exportText(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let path = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let text = doc.getText()
        try text.write(toFile: path, atomically: true, encoding: .utf8)

        return "Exported text to: \(path)"
    }

    private func exportMarkdown(args: [String: Value]) async throws -> String {
        guard let docId = args["doc_id"]?.stringValue else {
            throw WordError.missingParameter("doc_id")
        }
        guard let path = args["path"]?.stringValue else {
            throw WordError.missingParameter("path")
        }
        guard let doc = openDocuments[docId] else {
            throw WordError.documentNotFound(docId)
        }

        let markdown = doc.toMarkdown()
        try markdown.write(toFile: path, atomically: true, encoding: .utf8)

        return "Exported Markdown to: \(path)"
    }
}
