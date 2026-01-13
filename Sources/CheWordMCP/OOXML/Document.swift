import Foundation

/// Word 文件結構
struct WordDocument {
    var body: Body
    var styles: [Style]
    var properties: DocumentProperties
    var numbering: Numbering
    var sectionProperties: SectionProperties
    var headers: [Header]
    var footers: [Footer]
    var images: [ImageReference]              // 圖片資源
    var hyperlinkReferences: [HyperlinkReference] = []  // 超連結關係（用於 .rels）
    var comments: CommentsCollection = CommentsCollection()   // 註解集合
    var revisions: RevisionsCollection = RevisionsCollection() // 修訂集合
    var footnotes: FootnotesCollection = FootnotesCollection() // 腳註集合
    var endnotes: EndnotesCollection = EndnotesCollection()    // 尾註集合
    private var nextBookmarkId: Int = 1       // 書籤 ID 計數器
    private var nextHyperlinkId: Int = 1      // 超連結 ID 計數器

    init() {
        self.body = Body()
        self.styles = Style.defaultStyles
        self.properties = DocumentProperties()
        self.numbering = Numbering()
        self.sectionProperties = SectionProperties()
        self.headers = []
        self.footers = []
        self.images = []
        self.hyperlinkReferences = []
        self.comments = CommentsCollection()
        self.revisions = RevisionsCollection()
        self.footnotes = FootnotesCollection()
        self.endnotes = EndnotesCollection()
    }

    // MARK: - Document Info

    struct Info {
        let paragraphCount: Int
        let characterCount: Int
        let wordCount: Int
        let tableCount: Int
    }

    func getInfo() -> Info {
        let paragraphs = getParagraphs()
        let text = getText()
        let words = text.split(whereSeparator: { $0.isWhitespace || $0.isNewline })

        return Info(
            paragraphCount: paragraphs.count,
            characterCount: text.count,
            wordCount: words.count,
            tableCount: body.tables.count
        )
    }

    // MARK: - Text Operations

    func getText() -> String {
        var result = ""
        for child in body.children {
            switch child {
            case .paragraph(let para):
                result += para.getText() + "\n"
            case .table(let table):
                result += table.getText() + "\n"
            }
        }
        return result.trimmingCharacters(in: .whitespacesAndNewlines)
    }

    func getParagraphs() -> [Paragraph] {
        return body.children.compactMap { child in
            if case .paragraph(let para) = child {
                return para
            }
            return nil
        }
    }

    // MARK: - Paragraph Operations

    mutating func appendParagraph(_ paragraph: Paragraph) {
        body.children.append(.paragraph(paragraph))
    }

    mutating func insertParagraph(_ paragraph: Paragraph, at index: Int) {
        let clampedIndex = min(max(0, index), body.children.count)
        body.children.insert(.paragraph(paragraph), at: clampedIndex)
    }

    mutating func updateParagraph(at index: Int, text: String) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.runs = [Run(text: text)]
            body.children[actualIndex] = .paragraph(para)
        }
    }

    mutating func deleteParagraph(at index: Int) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        body.children.remove(at: actualIndex)
    }

    mutating func replaceText(find: String, with replacement: String, all: Bool) -> Int {
        var count = 0
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                for j in 0..<para.runs.count {
                    if para.runs[j].text.contains(find) {
                        if all {
                            let occurrences = para.runs[j].text.components(separatedBy: find).count - 1
                            count += occurrences
                            para.runs[j].text = para.runs[j].text.replacingOccurrences(of: find, with: replacement)
                        } else if count == 0 {
                            if let range = para.runs[j].text.range(of: find) {
                                para.runs[j].text.replaceSubrange(range, with: replacement)
                                count = 1
                            }
                        }
                    }
                }
                body.children[i] = .paragraph(para)
                if !all && count > 0 { break }
            }
        }
        return count
    }

    // MARK: - Formatting

    mutating func formatParagraph(at index: Int, with format: RunProperties) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            for i in 0..<para.runs.count {
                para.runs[i].properties.merge(with: format)
            }
            body.children[actualIndex] = .paragraph(para)
        }
    }

    mutating func setParagraphFormat(at index: Int, properties: ParagraphProperties) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.merge(with: properties)
            body.children[actualIndex] = .paragraph(para)
        }
    }

    mutating func applyStyle(at index: Int, style: String) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.style = style
            body.children[actualIndex] = .paragraph(para)
        }
    }

    // MARK: - Table Operations

    mutating func appendTable(_ table: Table) {
        body.children.append(.table(table))
        body.tables.append(table)
    }

    mutating func insertTable(_ table: Table, at index: Int) {
        let clampedIndex = min(max(0, index), body.children.count)
        body.children.insert(.table(table), at: clampedIndex)
        body.tables.append(table)
    }

    /// 取得所有表格
    func getTables() -> [Table] {
        return body.children.compactMap { child in
            if case .table(let table) = child {
                return table
            }
            return nil
        }
    }

    /// 取得表格索引對應到 body.children 的實際索引
    private func getTableIndices() -> [Int] {
        return body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .table = child { return i }
            return nil
        }
    }

    /// 更新表格儲存格內容
    mutating func updateCell(tableIndex: Int, row: Int, col: Int, text: String) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            guard row >= 0 && row < table.rows.count else {
                throw WordError.invalidIndex(row)
            }
            guard col >= 0 && col < table.rows[row].cells.count else {
                throw WordError.invalidIndex(col)
            }

            table.rows[row].cells[col] = TableCell(text: text)
            body.children[actualIndex] = .table(table)

            // 同步更新 body.tables
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
    }

    /// 刪除表格
    mutating func deleteTable(at tableIndex: Int) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        body.children.remove(at: actualIndex)

        // 同步更新 body.tables
        if tableIndex < body.tables.count {
            body.tables.remove(at: tableIndex)
        }
    }

    /// 合併儲存格（水平合併）
    mutating func mergeCellsHorizontal(tableIndex: Int, row: Int, startCol: Int, endCol: Int) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            guard row >= 0 && row < table.rows.count else {
                throw WordError.invalidIndex(row)
            }
            guard startCol >= 0 && startCol < table.rows[row].cells.count else {
                throw WordError.invalidIndex(startCol)
            }
            guard endCol >= startCol && endCol < table.rows[row].cells.count else {
                throw WordError.invalidIndex(endCol)
            }

            // 設定第一個儲存格的 gridSpan
            let span = endCol - startCol + 1
            table.rows[row].cells[startCol].properties.gridSpan = span

            // 移除被合併的儲存格（從後往前移除以保持索引正確）
            for col in (startCol + 1...endCol).reversed() {
                table.rows[row].cells.remove(at: col)
            }

            body.children[actualIndex] = .table(table)
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
    }

    /// 合併儲存格（垂直合併）
    mutating func mergeCellsVertical(tableIndex: Int, col: Int, startRow: Int, endRow: Int) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            guard startRow >= 0 && startRow < table.rows.count else {
                throw WordError.invalidIndex(startRow)
            }
            guard endRow >= startRow && endRow < table.rows.count else {
                throw WordError.invalidIndex(endRow)
            }

            // 設定第一個儲存格為 restart
            if col < table.rows[startRow].cells.count {
                table.rows[startRow].cells[col].properties.verticalMerge = .restart
            }

            // 設定其餘儲存格為 continue
            for row in (startRow + 1)...endRow {
                if col < table.rows[row].cells.count {
                    table.rows[row].cells[col].properties.verticalMerge = .continue
                }
            }

            body.children[actualIndex] = .table(table)
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
    }

    /// 設定表格樣式（邊框）
    mutating func setTableBorders(tableIndex: Int, borders: TableBorders) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            table.properties.borders = borders
            body.children[actualIndex] = .table(table)
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
    }

    /// 設定儲存格底色
    mutating func setCellShading(tableIndex: Int, row: Int, col: Int, shading: CellShading) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            guard row >= 0 && row < table.rows.count else {
                throw WordError.invalidIndex(row)
            }
            guard col >= 0 && col < table.rows[row].cells.count else {
                throw WordError.invalidIndex(col)
            }

            table.rows[row].cells[col].properties.shading = shading
            body.children[actualIndex] = .table(table)
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
    }

    // MARK: - Style Management

    /// 取得所有樣式
    func getStyles() -> [Style] {
        return styles
    }

    /// 根據 ID 取得樣式
    func getStyle(by id: String) -> Style? {
        return styles.first { $0.id == id }
    }

    /// 新增自訂樣式
    mutating func addStyle(_ style: Style) throws {
        // 檢查是否已存在相同 ID 的樣式
        if styles.contains(where: { $0.id == style.id }) {
            throw WordError.invalidFormat("Style with id '\(style.id)' already exists")
        }
        styles.append(style)
    }

    /// 更新樣式
    mutating func updateStyle(id: String, with updates: StyleUpdate) throws {
        guard let index = styles.firstIndex(where: { $0.id == id }) else {
            throw WordError.invalidFormat("Style '\(id)' not found")
        }

        var style = styles[index]

        if let name = updates.name {
            style.name = name
        }
        if let basedOn = updates.basedOn {
            style.basedOn = basedOn
        }
        if let nextStyle = updates.nextStyle {
            style.nextStyle = nextStyle
        }
        if let isQuickStyle = updates.isQuickStyle {
            style.isQuickStyle = isQuickStyle
        }
        if let paragraphProps = updates.paragraphProperties {
            if style.paragraphProperties == nil {
                style.paragraphProperties = ParagraphProperties()
            }
            style.paragraphProperties?.merge(with: paragraphProps)
        }
        if let runProps = updates.runProperties {
            if style.runProperties == nil {
                style.runProperties = RunProperties()
            }
            style.runProperties?.merge(with: runProps)
        }

        styles[index] = style
    }

    /// 刪除樣式（不能刪除預設樣式）
    mutating func deleteStyle(id: String) throws {
        guard let index = styles.firstIndex(where: { $0.id == id }) else {
            throw WordError.invalidFormat("Style '\(id)' not found")
        }

        // 不能刪除預設樣式
        if styles[index].isDefault {
            throw WordError.invalidFormat("Cannot delete default style '\(id)'")
        }

        // 檢查是否為內建樣式（Normal, Heading1-3, Title, Subtitle）
        let builtInIds = ["Normal", "Heading1", "Heading2", "Heading3", "Title", "Subtitle"]
        if builtInIds.contains(id) {
            throw WordError.invalidFormat("Cannot delete built-in style '\(id)'")
        }

        styles.remove(at: index)
    }

    // MARK: - List Operations

    /// 插入項目符號清單
    mutating func insertBulletList(items: [String], at index: Int? = nil) -> Int {
        let numId = numbering.createBulletList()

        for (itemIndex, text) in items.enumerated() {
            var para = Paragraph(text: text)
            para.properties.numbering = NumberingInfo(numId: numId, level: 0)

            if let index = index {
                insertParagraph(para, at: index + itemIndex)
            } else {
                appendParagraph(para)
            }
        }

        return numId
    }

    /// 插入編號清單
    mutating func insertNumberedList(items: [String], at index: Int? = nil) -> Int {
        let numId = numbering.createNumberedList()

        for (itemIndex, text) in items.enumerated() {
            var para = Paragraph(text: text)
            para.properties.numbering = NumberingInfo(numId: numId, level: 0)

            if let index = index {
                insertParagraph(para, at: index + itemIndex)
            } else {
                appendParagraph(para)
            }
        }

        return numId
    }

    /// 設定段落的清單層級
    mutating func setListLevel(paragraphIndex: Int, level: Int) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            guard para.properties.numbering != nil else {
                throw WordError.invalidFormat("Paragraph is not part of a list")
            }

            guard level >= 0 && level <= 8 else {
                throw WordError.invalidParameter("level", "Must be between 0 and 8")
            }

            para.properties.numbering?.level = level
            body.children[actualIndex] = .paragraph(para)
        }
    }

    /// 將段落添加到現有清單
    mutating func addToList(paragraphIndex: Int, numId: Int, level: Int = 0) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.numbering = NumberingInfo(numId: numId, level: level)
            body.children[actualIndex] = .paragraph(para)
        }
    }

    /// 移除段落的清單格式
    mutating func removeFromList(paragraphIndex: Int) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.numbering = nil
            body.children[actualIndex] = .paragraph(para)
        }
    }

    // MARK: - Page Settings

    /// 設定頁面大小
    mutating func setPageSize(_ size: PageSize) {
        sectionProperties.pageSize = size
    }

    /// 設定頁面大小（使用名稱）
    mutating func setPageSize(name: String) throws {
        guard let size = PageSize.from(name: name) else {
            throw WordError.invalidParameter("pageSize", "Unknown page size: \(name). Valid options: letter, a4, legal, a3, a5, b5, executive")
        }
        sectionProperties.pageSize = size
    }

    /// 設定頁邊距
    mutating func setPageMargins(_ margins: PageMargins) {
        sectionProperties.pageMargins = margins
    }

    /// 設定頁邊距（使用名稱）
    mutating func setPageMargins(name: String) throws {
        guard let margins = PageMargins.from(name: name) else {
            throw WordError.invalidParameter("margins", "Unknown margin preset: \(name). Valid options: normal, narrow, moderate, wide")
        }
        sectionProperties.pageMargins = margins
    }

    /// 設定頁邊距（使用具體數值，單位：twips）
    mutating func setPageMargins(top: Int? = nil, right: Int? = nil, bottom: Int? = nil, left: Int? = nil) {
        if let top = top {
            sectionProperties.pageMargins.top = top
        }
        if let right = right {
            sectionProperties.pageMargins.right = right
        }
        if let bottom = bottom {
            sectionProperties.pageMargins.bottom = bottom
        }
        if let left = left {
            sectionProperties.pageMargins.left = left
        }
    }

    /// 設定頁面方向
    mutating func setPageOrientation(_ orientation: PageOrientation) {
        sectionProperties.orientation = orientation

        // 如果切換方向，也要交換頁面寬高
        if orientation == .landscape && sectionProperties.pageSize.width < sectionProperties.pageSize.height {
            sectionProperties.pageSize = sectionProperties.pageSize.landscape
        } else if orientation == .portrait && sectionProperties.pageSize.width > sectionProperties.pageSize.height {
            sectionProperties.pageSize = sectionProperties.pageSize.landscape
        }
    }

    /// 插入分頁符
    mutating func insertPageBreak(at paragraphIndex: Int? = nil) {
        // 分頁符是一個特殊的段落，只包含 <w:br w:type="page"/>
        var para = Paragraph()
        para.hasPageBreak = true

        if let index = paragraphIndex {
            insertParagraph(para, at: index)
        } else {
            appendParagraph(para)
        }
    }

    /// 插入分節符
    mutating func insertSectionBreak(type: SectionBreakType = .nextPage, at paragraphIndex: Int? = nil) {
        // 分節符放在段落屬性中
        var para = Paragraph()
        para.properties.sectionBreak = type

        if let index = paragraphIndex {
            insertParagraph(para, at: index)
        } else {
            appendParagraph(para)
        }
    }

    // MARK: - Header/Footer Operations

    /// 取得下一個可用的關係 ID
    private var nextRelationshipId: String {
        // 基本 ID 從 rId4 開始（rId1-3 用於基本關係）
        // 加上 numbering 的話從 rId5 開始
        let baseId = numbering.abstractNums.isEmpty ? 4 : 5
        let usedCount = headers.count + footers.count
        return "rId\(baseId + usedCount)"
    }

    /// 新增頁首
    mutating func addHeader(text: String, type: HeaderFooterType = .default) -> Header {
        let id = nextRelationshipId
        let header = Header.withText(text, id: id, type: type)
        headers.append(header)

        // 更新分節屬性中的頁首參照
        if type == .default {
            sectionProperties.headerReference = id
        }

        return header
    }

    /// 新增含頁碼的頁首
    mutating func addHeaderWithPageNumber(type: HeaderFooterType = .default) -> Header {
        let id = nextRelationshipId
        let header = Header.withPageNumber(id: id, type: type)
        headers.append(header)

        if type == .default {
            sectionProperties.headerReference = id
        }

        return header
    }

    /// 更新頁首內容
    mutating func updateHeader(id: String, text: String) throws {
        guard let index = headers.firstIndex(where: { $0.id == id }) else {
            throw WordError.invalidFormat("Header '\(id)' not found")
        }

        var header = headers[index]
        header.paragraphs = [Paragraph(text: text)]
        headers[index] = header
    }

    /// 新增頁尾
    mutating func addFooter(text: String, type: HeaderFooterType = .default) -> Footer {
        let id = nextRelationshipId
        let footer = Footer.withText(text, id: id, type: type)
        footers.append(footer)

        // 更新分節屬性中的頁尾參照
        if type == .default {
            sectionProperties.footerReference = id
        }

        return footer
    }

    /// 新增含頁碼的頁尾
    mutating func addFooterWithPageNumber(format: PageNumberFormat = .simple, type: HeaderFooterType = .default) -> Footer {
        let id = nextRelationshipId
        let footer = Footer.withPageNumber(id: id, format: format, type: type)
        footers.append(footer)

        if type == .default {
            sectionProperties.footerReference = id
        }

        return footer
    }

    /// 更新頁尾內容
    mutating func updateFooter(id: String, text: String) throws {
        guard let index = footers.firstIndex(where: { $0.id == id }) else {
            throw WordError.invalidFormat("Footer '\(id)' not found")
        }

        var footer = footers[index]
        footer.paragraphs = [Paragraph(text: text)]
        footers[index] = footer
    }

    // MARK: - Image Operations

    /// 取得下一個可用的圖片關係 ID
    private var nextImageRelationshipId: String {
        // 基本 ID 從 rId4 開始（rId1-3 用於基本關係）
        // 加上 numbering 的話從 rId5 開始
        let baseId = numbering.abstractNums.isEmpty ? 4 : 5
        let usedCount = headers.count + footers.count + images.count
        return "rId\(baseId + usedCount)"
    }

    /// 從 Base64 插入圖片
    mutating func insertImage(
        base64: String,
        fileName: String,
        widthPx: Int,
        heightPx: Int,
        at paragraphIndex: Int? = nil,
        name: String = "Picture",
        description: String = ""
    ) throws -> String {
        let imageId = nextImageRelationshipId

        // 建立圖片參照
        let imageRef = try ImageReference.from(base64: base64, fileName: fileName, id: imageId)
        images.append(imageRef)

        // 建立 Drawing
        var drawing = Drawing.from(widthPx: widthPx, heightPx: heightPx, imageId: imageId, name: name)
        drawing.description = description

        // 建立含圖片的 Run
        let run = Run.withDrawing(drawing)

        // 建立段落
        let para = Paragraph(runs: [run])

        // 插入段落
        if let index = paragraphIndex {
            insertParagraph(para, at: index)
        } else {
            appendParagraph(para)
        }

        return imageId
    }

    /// 從檔案路徑插入圖片
    mutating func insertImage(
        path: String,
        widthPx: Int,
        heightPx: Int,
        at paragraphIndex: Int? = nil,
        name: String = "Picture",
        description: String = ""
    ) throws -> String {
        let imageId = nextImageRelationshipId

        // 建立圖片參照
        let imageRef = try ImageReference.from(path: path, id: imageId)
        images.append(imageRef)

        // 建立 Drawing
        var drawing = Drawing.from(widthPx: widthPx, heightPx: heightPx, imageId: imageId, name: name)
        drawing.description = description

        // 建立含圖片的 Run
        let run = Run.withDrawing(drawing)

        // 建立段落
        let para = Paragraph(runs: [run])

        // 插入段落
        if let index = paragraphIndex {
            insertParagraph(para, at: index)
        } else {
            appendParagraph(para)
        }

        return imageId
    }

    /// 更新圖片大小
    mutating func updateImage(imageId: String, widthPx: Int? = nil, heightPx: Int? = nil) throws {
        // 搜尋所有段落找到含有此圖片的 Run
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                for j in 0..<para.runs.count {
                    if var drawing = para.runs[j].drawing, drawing.imageId == imageId {
                        if let w = widthPx {
                            drawing.width = w * 9525  // 轉換為 EMU
                        }
                        if let h = heightPx {
                            drawing.height = h * 9525
                        }
                        para.runs[j].drawing = drawing
                        body.children[i] = .paragraph(para)
                        return
                    }
                }
            }
        }
        throw WordError.invalidFormat("Image '\(imageId)' not found")
    }

    /// 設定圖片樣式
    mutating func setImageStyle(
        imageId: String,
        hasBorder: Bool? = nil,
        borderColor: String? = nil,
        borderWidth: Int? = nil,
        hasShadow: Bool? = nil
    ) throws {
        // 搜尋所有段落找到含有此圖片的 Run
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                for j in 0..<para.runs.count {
                    if var drawing = para.runs[j].drawing, drawing.imageId == imageId {
                        if let border = hasBorder {
                            drawing.hasBorder = border
                        }
                        if let color = borderColor {
                            drawing.borderColor = color
                        }
                        if let width = borderWidth {
                            drawing.borderWidth = width
                        }
                        if let shadow = hasShadow {
                            drawing.hasShadow = shadow
                        }
                        para.runs[j].drawing = drawing
                        body.children[i] = .paragraph(para)
                        return
                    }
                }
            }
        }
        throw WordError.invalidFormat("Image '\(imageId)' not found")
    }

    /// 刪除圖片
    mutating func deleteImage(imageId: String) throws {
        // 移除圖片資源
        guard let resourceIndex = images.firstIndex(where: { $0.id == imageId }) else {
            throw WordError.invalidFormat("Image '\(imageId)' not found")
        }
        images.remove(at: resourceIndex)

        // 搜尋並移除含有此圖片的 Run
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                para.runs.removeAll { $0.drawing?.imageId == imageId }
                body.children[i] = .paragraph(para)
            }
        }
    }

    /// 取得所有圖片資訊
    func getImages() -> [(id: String, fileName: String, widthPx: Int, heightPx: Int)] {
        var result: [(id: String, fileName: String, widthPx: Int, heightPx: Int)] = []

        // 從 images 和 body 中收集圖片資訊
        for image in images {
            // 找對應的 Drawing 來取得尺寸
            var widthPx = 0
            var heightPx = 0

            for child in body.children {
                if case .paragraph(let para) = child {
                    for run in para.runs {
                        if let drawing = run.drawing, drawing.imageId == image.id {
                            widthPx = drawing.widthInPixels
                            heightPx = drawing.heightInPixels
                            break
                        }
                    }
                }
            }

            result.append((id: image.id, fileName: image.fileName, widthPx: widthPx, heightPx: heightPx))
        }

        return result
    }

    // MARK: - Hyperlink Operations

    /// 取得下一個可用的超連結關係 ID
    private var nextHyperlinkRelationshipId: String {
        // 基本 ID 從 rId4 開始
        let baseId = numbering.abstractNums.isEmpty ? 4 : 5
        let usedCount = headers.count + footers.count + images.count + hyperlinkReferences.count
        return "rId\(baseId + usedCount)"
    }

    /// 插入外部超連結
    mutating func insertHyperlink(
        url: String,
        text: String,
        at paragraphIndex: Int? = nil,
        tooltip: String? = nil
    ) -> String {
        let hyperlinkId = "hyperlink_\(nextHyperlinkId)"
        nextHyperlinkId += 1

        let relationshipId = nextHyperlinkRelationshipId

        // 建立超連結關係（用於 .rels 檔案）
        let reference = HyperlinkReference(relationshipId: relationshipId, url: url)
        hyperlinkReferences.append(reference)

        // 建立超連結
        let hyperlink = Hyperlink.external(
            id: hyperlinkId,
            text: text,
            url: url,
            relationshipId: relationshipId,
            tooltip: tooltip
        )

        // 如果指定了段落索引，加到該段落；否則建立新段落
        if let index = paragraphIndex {
            let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            if index >= 0 && index < paragraphIndices.count {
                let actualIndex = paragraphIndices[index]
                if case .paragraph(var para) = body.children[actualIndex] {
                    para.hyperlinks.append(hyperlink)
                    body.children[actualIndex] = .paragraph(para)
                }
            }
        } else {
            // 建立新段落包含超連結
            var para = Paragraph()
            para.hyperlinks.append(hyperlink)
            appendParagraph(para)
        }

        return hyperlinkId
    }

    /// 插入內部連結（連結到書籤）
    mutating func insertInternalLink(
        bookmarkName: String,
        text: String,
        at paragraphIndex: Int? = nil,
        tooltip: String? = nil
    ) -> String {
        let hyperlinkId = "hyperlink_\(nextHyperlinkId)"
        nextHyperlinkId += 1

        // 內部連結不需要 relationship
        let hyperlink = Hyperlink.internal(
            id: hyperlinkId,
            text: text,
            bookmarkName: bookmarkName,
            tooltip: tooltip
        )

        if let index = paragraphIndex {
            let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            if index >= 0 && index < paragraphIndices.count {
                let actualIndex = paragraphIndices[index]
                if case .paragraph(var para) = body.children[actualIndex] {
                    para.hyperlinks.append(hyperlink)
                    body.children[actualIndex] = .paragraph(para)
                }
            }
        } else {
            var para = Paragraph()
            para.hyperlinks.append(hyperlink)
            appendParagraph(para)
        }

        return hyperlinkId
    }

    /// 更新超連結
    mutating func updateHyperlink(hyperlinkId: String, text: String? = nil, url: String? = nil) throws {
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                for j in 0..<para.hyperlinks.count {
                    if para.hyperlinks[j].id == hyperlinkId {
                        if let newText = text {
                            para.hyperlinks[j].text = newText
                        }
                        if let newUrl = url {
                            // 更新 URL 需要同時更新 relationship
                            if let rId = para.hyperlinks[j].relationshipId {
                                if let refIndex = hyperlinkReferences.firstIndex(where: { $0.relationshipId == rId }) {
                                    hyperlinkReferences[refIndex].url = newUrl
                                }
                            }
                            para.hyperlinks[j].url = newUrl
                        }
                        body.children[i] = .paragraph(para)
                        return
                    }
                }
            }
        }
        throw WordError.invalidFormat("Hyperlink '\(hyperlinkId)' not found")
    }

    /// 刪除超連結
    mutating func deleteHyperlink(hyperlinkId: String) throws {
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                if let index = para.hyperlinks.firstIndex(where: { $0.id == hyperlinkId }) {
                    // 如果是外部連結，也要刪除 relationship
                    if let rId = para.hyperlinks[index].relationshipId {
                        hyperlinkReferences.removeAll { $0.relationshipId == rId }
                    }
                    para.hyperlinks.remove(at: index)
                    body.children[i] = .paragraph(para)
                    return
                }
            }
        }
        throw WordError.invalidFormat("Hyperlink '\(hyperlinkId)' not found")
    }

    /// 列出所有超連結
    func getHyperlinks() -> [(id: String, text: String, url: String?, anchor: String?, type: String)] {
        var result: [(id: String, text: String, url: String?, anchor: String?, type: String)] = []

        for child in body.children {
            if case .paragraph(let para) = child {
                for hyperlink in para.hyperlinks {
                    let typeStr = hyperlink.type == .external ? "external" : "internal"
                    result.append((
                        id: hyperlink.id,
                        text: hyperlink.text,
                        url: hyperlink.url,
                        anchor: hyperlink.anchor,
                        type: typeStr
                    ))
                }
            }
        }

        return result
    }

    // MARK: - Bookmark Operations

    /// 插入書籤
    mutating func insertBookmark(
        name: String,
        at paragraphIndex: Int? = nil
    ) throws -> Int {
        // 驗證書籤名稱
        let normalizedName = Bookmark.normalizeName(name)
        guard Bookmark.validateName(normalizedName) else {
            throw BookmarkError.invalidName(name)
        }

        // 檢查是否已存在同名書籤
        for child in body.children {
            if case .paragraph(let para) = child {
                if para.bookmarks.contains(where: { $0.name == normalizedName }) {
                    throw BookmarkError.duplicateName(normalizedName)
                }
            }
        }

        let bookmarkId = nextBookmarkId
        nextBookmarkId += 1

        let bookmark = Bookmark(id: bookmarkId, name: normalizedName)

        if let index = paragraphIndex {
            let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            guard index >= 0 && index < paragraphIndices.count else {
                throw WordError.invalidIndex(index)
            }

            let actualIndex = paragraphIndices[index]
            if case .paragraph(var para) = body.children[actualIndex] {
                para.bookmarks.append(bookmark)
                body.children[actualIndex] = .paragraph(para)
            }
        } else {
            // 加到最後一個段落，如果沒有段落則建立新的
            if let lastIndex = body.children.lastIndex(where: {
                if case .paragraph = $0 { return true }
                return false
            }) {
                if case .paragraph(var para) = body.children[lastIndex] {
                    para.bookmarks.append(bookmark)
                    body.children[lastIndex] = .paragraph(para)
                }
            } else {
                var para = Paragraph()
                para.bookmarks.append(bookmark)
                appendParagraph(para)
            }
        }

        return bookmarkId
    }

    /// 刪除書籤
    mutating func deleteBookmark(name: String) throws {
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                if let index = para.bookmarks.firstIndex(where: { $0.name == name }) {
                    para.bookmarks.remove(at: index)
                    body.children[i] = .paragraph(para)
                    return
                }
            }
        }
        throw BookmarkError.notFound(name)
    }

    /// 列出所有書籤
    func getBookmarks() -> [(id: Int, name: String, paragraphIndex: Int)] {
        var result: [(id: Int, name: String, paragraphIndex: Int)] = []
        var paragraphCount = 0

        for child in body.children {
            if case .paragraph(let para) = child {
                for bookmark in para.bookmarks {
                    result.append((id: bookmark.id, name: bookmark.name, paragraphIndex: paragraphCount))
                }
                paragraphCount += 1
            }
        }

        return result
    }

    // MARK: - Comment Operations

    /// 插入註解
    mutating func insertComment(
        text: String,
        author: String,
        paragraphIndex: Int
    ) throws -> Int {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw CommentError.invalidParagraphIndex(paragraphIndex)
        }

        let commentId = comments.nextCommentId()
        let comment = Comment(
            id: commentId,
            author: author,
            text: text,
            paragraphIndex: paragraphIndex
        )

        comments.comments.append(comment)

        // 在段落中添加註解標記
        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.commentIds.append(commentId)
            body.children[actualIndex] = .paragraph(para)
        }

        return commentId
    }

    /// 更新註解
    mutating func updateComment(commentId: Int, text: String) throws {
        guard let index = comments.comments.firstIndex(where: { $0.id == commentId }) else {
            throw CommentError.notFound(commentId)
        }

        comments.comments[index].text = text
    }

    /// 刪除註解
    mutating func deleteComment(commentId: Int) throws {
        guard let index = comments.comments.firstIndex(where: { $0.id == commentId }) else {
            throw CommentError.notFound(commentId)
        }

        // 從段落中移除註解標記
        let paragraphIndex = comments.comments[index].paragraphIndex
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        if paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count {
            let actualIndex = paragraphIndices[paragraphIndex]
            if case .paragraph(var para) = body.children[actualIndex] {
                para.commentIds.removeAll { $0 == commentId }
                body.children[actualIndex] = .paragraph(para)
            }
        }

        comments.comments.remove(at: index)
    }

    /// 列出所有註解
    func getComments() -> [(id: Int, author: String, text: String, paragraphIndex: Int, date: Date)] {
        return comments.comments.map { comment in
            (id: comment.id, author: comment.author, text: comment.text,
             paragraphIndex: comment.paragraphIndex, date: comment.date)
        }
    }

    // MARK: - Track Changes Operations

    /// 啟用修訂追蹤
    mutating func enableTrackChanges(author: String = "Unknown") {
        revisions.settings.enabled = true
        revisions.settings.author = author
        revisions.settings.dateTime = Date()
    }

    /// 停用修訂追蹤
    mutating func disableTrackChanges() {
        revisions.settings.enabled = false
    }

    /// 檢查修訂追蹤是否啟用
    func isTrackChangesEnabled() -> Bool {
        return revisions.settings.enabled
    }

    /// 取得所有修訂
    func getRevisions() -> [(id: Int, type: String, author: String, paragraphIndex: Int, originalText: String?, newText: String?)] {
        return revisions.revisions.map { rev in
            (id: rev.id, type: rev.type.rawValue, author: rev.author,
             paragraphIndex: rev.paragraphIndex, originalText: rev.originalText, newText: rev.newText)
        }
    }

    /// 接受修訂
    mutating func acceptRevision(revisionId: Int) throws {
        guard let index = revisions.revisions.firstIndex(where: { $0.id == revisionId }) else {
            throw RevisionError.notFound(revisionId)
        }

        let revision = revisions.revisions[index]

        // 根據修訂類型處理
        switch revision.type {
        case .insertion:
            // 接受插入：移除標記，保留文字（文字已在文件中）
            break
        case .deletion:
            // 接受刪除：實際移除被標記為刪除的文字
            let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            if revision.paragraphIndex >= 0 && revision.paragraphIndex < paragraphIndices.count {
                let actualIndex = paragraphIndices[revision.paragraphIndex]
                if case .paragraph(var para) = body.children[actualIndex] {
                    // 移除被刪除的文字
                    if let originalText = revision.originalText {
                        for j in 0..<para.runs.count {
                            para.runs[j].text = para.runs[j].text.replacingOccurrences(of: originalText, with: "")
                        }
                    }
                    body.children[actualIndex] = .paragraph(para)
                }
            }
        case .formatting, .paragraphChange:
            // 接受格式變更：保留新格式
            break
        }

        // 移除修訂記錄
        revisions.revisions.remove(at: index)
    }

    /// 拒絕修訂
    mutating func rejectRevision(revisionId: Int) throws {
        guard let index = revisions.revisions.firstIndex(where: { $0.id == revisionId }) else {
            throw RevisionError.notFound(revisionId)
        }

        let revision = revisions.revisions[index]

        // 根據修訂類型處理
        switch revision.type {
        case .insertion:
            // 拒絕插入：移除插入的文字
            let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            if revision.paragraphIndex >= 0 && revision.paragraphIndex < paragraphIndices.count {
                let actualIndex = paragraphIndices[revision.paragraphIndex]
                if case .paragraph(var para) = body.children[actualIndex] {
                    // 移除插入的文字
                    if let newText = revision.newText {
                        for j in 0..<para.runs.count {
                            para.runs[j].text = para.runs[j].text.replacingOccurrences(of: newText, with: "")
                        }
                    }
                    body.children[actualIndex] = .paragraph(para)
                }
            }
        case .deletion:
            // 拒絕刪除：恢復被刪除的文字（文字已在標記中，移除標記即可）
            break
        case .formatting, .paragraphChange:
            // 拒絕格式變更：恢復原格式（需要實作格式恢復邏輯）
            break
        }

        // 移除修訂記錄
        revisions.revisions.remove(at: index)
    }

    /// 接受所有修訂
    mutating func acceptAllRevisions() {
        // 從後往前接受，避免索引問題
        for revision in revisions.revisions.reversed() {
            try? acceptRevision(revisionId: revision.id)
        }
    }

    /// 拒絕所有修訂
    mutating func rejectAllRevisions() {
        // 從後往前拒絕，避免索引問題
        for revision in revisions.revisions.reversed() {
            try? rejectRevision(revisionId: revision.id)
        }
    }

    // MARK: - Footnote Operations

    /// 插入腳註
    mutating func insertFootnote(
        text: String,
        paragraphIndex: Int
    ) throws -> Int {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw FootnoteError.invalidParagraphIndex(paragraphIndex)
        }

        let footnoteId = footnotes.nextFootnoteId()
        let footnote = Footnote(
            id: footnoteId,
            text: text,
            paragraphIndex: paragraphIndex
        )

        footnotes.footnotes.append(footnote)

        // 在段落中添加腳註參照
        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.footnoteIds.append(footnoteId)
            body.children[actualIndex] = .paragraph(para)
        }

        return footnoteId
    }

    /// 刪除腳註
    mutating func deleteFootnote(footnoteId: Int) throws {
        guard let index = footnotes.footnotes.firstIndex(where: { $0.id == footnoteId }) else {
            throw FootnoteError.notFound(footnoteId)
        }

        // 從段落中移除腳註參照
        let paragraphIndex = footnotes.footnotes[index].paragraphIndex
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        if paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count {
            let actualIndex = paragraphIndices[paragraphIndex]
            if case .paragraph(var para) = body.children[actualIndex] {
                para.footnoteIds.removeAll { $0 == footnoteId }
                body.children[actualIndex] = .paragraph(para)
            }
        }

        footnotes.footnotes.remove(at: index)
    }

    /// 列出所有腳註
    func getFootnotes() -> [(id: Int, text: String, paragraphIndex: Int)] {
        return footnotes.footnotes.map { footnote in
            (id: footnote.id, text: footnote.text, paragraphIndex: footnote.paragraphIndex)
        }
    }

    // MARK: - Endnote Operations

    /// 插入尾註
    mutating func insertEndnote(
        text: String,
        paragraphIndex: Int
    ) throws -> Int {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw EndnoteError.invalidParagraphIndex(paragraphIndex)
        }

        let endnoteId = endnotes.nextEndnoteId()
        let endnote = Endnote(
            id: endnoteId,
            text: text,
            paragraphIndex: paragraphIndex
        )

        endnotes.endnotes.append(endnote)

        // 在段落中添加尾註參照
        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.endnoteIds.append(endnoteId)
            body.children[actualIndex] = .paragraph(para)
        }

        return endnoteId
    }

    /// 刪除尾註
    mutating func deleteEndnote(endnoteId: Int) throws {
        guard let index = endnotes.endnotes.firstIndex(where: { $0.id == endnoteId }) else {
            throw EndnoteError.notFound(endnoteId)
        }

        // 從段落中移除尾註參照
        let paragraphIndex = endnotes.endnotes[index].paragraphIndex
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        if paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count {
            let actualIndex = paragraphIndices[paragraphIndex]
            if case .paragraph(var para) = body.children[actualIndex] {
                para.endnoteIds.removeAll { $0 == endnoteId }
                body.children[actualIndex] = .paragraph(para)
            }
        }

        endnotes.endnotes.remove(at: index)
    }

    /// 列出所有尾註
    func getEndnotes() -> [(id: Int, text: String, paragraphIndex: Int)] {
        return endnotes.endnotes.map { endnote in
            (id: endnote.id, text: endnote.text, paragraphIndex: endnote.paragraphIndex)
        }
    }

    // MARK: - Export

    func toMarkdown() -> String {
        var result = ""
        for child in body.children {
            switch child {
            case .paragraph(let para):
                let text = para.getText()
                if let style = para.properties.style {
                    switch style {
                    case "Heading1", "heading 1":
                        result += "# \(text)\n\n"
                    case "Heading2", "heading 2":
                        result += "## \(text)\n\n"
                    case "Heading3", "heading 3":
                        result += "### \(text)\n\n"
                    case "Title":
                        result += "# \(text)\n\n"
                    default:
                        result += "\(text)\n\n"
                    }
                } else {
                    result += "\(text)\n\n"
                }
            case .table(let table):
                result += table.toMarkdown() + "\n\n"
            }
        }
        return result.trimmingCharacters(in: .whitespacesAndNewlines)
    }
}

// MARK: - Body

struct Body {
    var children: [BodyChild] = []
    var tables: [Table] = []
}

enum BodyChild {
    case paragraph(Paragraph)
    case table(Table)
}

// MARK: - Document Properties

struct DocumentProperties {
    var title: String?
    var subject: String?
    var creator: String?
    var keywords: String?
    var description: String?
    var lastModifiedBy: String?
    var revision: Int?
    var created: Date?
    var modified: Date?
}

// MARK: - Advanced Features Extensions

extension WordDocument {
    // MARK: - Table of Contents

    /// 插入目錄
    mutating func insertTableOfContents(
        at index: Int? = nil,
        title: String? = "Contents",
        headingLevels: ClosedRange<Int> = 1...3,
        includePageNumbers: Bool = true,
        useHyperlinks: Bool = true
    ) {
        let toc = TableOfContents(
            title: title,
            headingLevels: headingLevels,
            includePageNumbers: includePageNumbers,
            useHyperlinks: useHyperlinks
        )

        // 建立包含 TOC XML 的段落
        var para = Paragraph()
        var run = Run(text: "")
        run.properties.rawXML = toc.toXML()
        para.runs = [run]

        if let idx = index {
            insertParagraph(para, at: idx)
        } else {
            // 預設插入到開頭
            body.children.insert(.paragraph(para), at: 0)
        }
    }

    // MARK: - Form Controls

    /// 插入文字欄位
    mutating func insertTextField(
        at paragraphIndex: Int,
        name: String,
        defaultValue: String? = nil,
        maxLength: Int? = nil
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let field = FormTextField(name: name, defaultValue: defaultValue, maxLength: maxLength)

        // 找到段落並添加欄位
        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph = child {
                if childIndex == paragraphIndex {
                    var run = Run(text: "")
                    run.properties.rawXML = field.toXML()
                    if case .paragraph(var para) = body.children[i] {
                        para.runs.append(run)
                        body.children[i] = .paragraph(para)
                    }
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 插入核取方塊
    mutating func insertCheckbox(
        at paragraphIndex: Int,
        name: String,
        isChecked: Bool = false
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let checkbox = FormCheckbox(name: name, isChecked: isChecked)

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph = child {
                if childIndex == paragraphIndex {
                    var run = Run(text: "")
                    run.properties.rawXML = checkbox.toXML()
                    if case .paragraph(var para) = body.children[i] {
                        para.runs.append(run)
                        body.children[i] = .paragraph(para)
                    }
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 插入下拉選單
    mutating func insertDropdown(
        at paragraphIndex: Int,
        name: String,
        options: [String],
        selectedIndex: Int = 0
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let dropdown = FormDropdown(name: name, options: options, selectedIndex: selectedIndex)

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph = child {
                if childIndex == paragraphIndex {
                    var run = Run(text: "")
                    run.properties.rawXML = dropdown.toXML()
                    if case .paragraph(var para) = body.children[i] {
                        para.runs.append(run)
                        body.children[i] = .paragraph(para)
                    }
                    return
                }
                childIndex += 1
            }
        }
    }

    // MARK: - Mathematical Equations

    /// 插入數學公式
    mutating func insertEquation(
        at paragraphIndex: Int? = nil,
        latex: String,
        displayMode: Bool = false
    ) {
        let equation = MathEquation(latex: latex, displayMode: displayMode)

        if displayMode {
            // 獨立區塊公式，建立新段落
            var para = Paragraph()
            var run = Run(text: "")
            run.properties.rawXML = equation.toXML()
            para.runs = [run]
            para.properties.alignment = .center

            if let idx = paragraphIndex {
                insertParagraph(para, at: idx)
            } else {
                appendParagraph(para)
            }
        } else {
            // 行內公式，加入到現有段落
            if let idx = paragraphIndex, idx >= 0 && idx < getParagraphs().count {
                var childIndex = 0
                for (i, child) in body.children.enumerated() {
                    if case .paragraph = child {
                        if childIndex == idx {
                            var run = Run(text: "")
                            run.properties.rawXML = equation.toXML()
                            if case .paragraph(var para) = body.children[i] {
                                para.runs.append(run)
                                body.children[i] = .paragraph(para)
                            }
                            return
                        }
                        childIndex += 1
                    }
                }
            }
        }
    }

    // MARK: - Advanced Paragraph Formatting

    /// 設定段落邊框
    mutating func setParagraphBorder(
        at paragraphIndex: Int,
        border: ParagraphBorder
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    para.properties.border = border
                    body.children[i] = .paragraph(para)
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 設定段落底色
    mutating func setParagraphShading(
        at paragraphIndex: Int,
        fill: String,
        pattern: ShadingPattern? = nil
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    para.properties.shading = CellShading(fill: fill, pattern: pattern)
                    body.children[i] = .paragraph(para)
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 設定字元間距
    mutating func setCharacterSpacing(
        at paragraphIndex: Int,
        spacing: Int? = nil,
        position: Int? = nil,
        kern: Int? = nil
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let charSpacing = CharacterSpacing(spacing: spacing, position: position, kern: kern)

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    for j in 0..<para.runs.count {
                        para.runs[j].properties.characterSpacing = charSpacing
                    }
                    body.children[i] = .paragraph(para)
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 設定文字效果
    mutating func setTextEffect(
        at paragraphIndex: Int,
        effect: TextEffect
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    for j in 0..<para.runs.count {
                        para.runs[j].properties.textEffect = effect
                    }
                    body.children[i] = .paragraph(para)
                    return
                }
                childIndex += 1
            }
        }
    }
}
