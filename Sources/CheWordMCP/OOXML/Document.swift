import Foundation

/// Word 文件結構
struct WordDocument {
    var body: Body
    var styles: [Style]
    var properties: DocumentProperties

    init() {
        self.body = Body()
        self.styles = Style.defaultStyles
        self.properties = DocumentProperties()
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
