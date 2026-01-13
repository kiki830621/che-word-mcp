import Foundation

// MARK: - Bookmark

/// 書籤（用於文件內部導航和連結）
struct Bookmark {
    var id: Int          // 書籤唯一 ID（整數）
    var name: String     // 書籤名稱（用於連結引用）

    init(id: Int, name: String) {
        self.id = id
        self.name = name
    }

    /// 驗證書籤名稱是否合法
    /// - 書籤名稱不能包含空格
    /// - 書籤名稱不能以數字開頭
    /// - 長度限制 40 字元
    static func validateName(_ name: String) -> Bool {
        guard !name.isEmpty else { return false }
        guard name.count <= 40 else { return false }
        guard !name.contains(" ") else { return false }
        guard let first = name.first, !first.isNumber else { return false }
        return true
    }

    /// 標準化書籤名稱（移除無效字元）
    static func normalizeName(_ name: String) -> String {
        var normalized = name
            .replacingOccurrences(of: " ", with: "_")
            .replacingOccurrences(of: "-", with: "_")

        // 如果以數字開頭，加上底線
        if let first = normalized.first, first.isNumber {
            normalized = "_" + normalized
        }

        // 限制長度
        if normalized.count > 40 {
            normalized = String(normalized.prefix(40))
        }

        return normalized
    }
}

// MARK: - Bookmark in Paragraph

/// 段落中的書籤資訊
struct ParagraphBookmark {
    var bookmark: Bookmark
    var startPosition: Int  // 在段落中的起始位置（字元索引）
    var endPosition: Int    // 在段落中的結束位置

    init(bookmark: Bookmark, startPosition: Int = 0, endPosition: Int = 0) {
        self.bookmark = bookmark
        self.startPosition = startPosition
        self.endPosition = endPosition
    }
}

// MARK: - XML Generation

extension Bookmark {
    /// 產生書籤開始標記
    func toBookmarkStartXML() -> String {
        return "<w:bookmarkStart w:id=\"\(id)\" w:name=\"\(escapeXML(name))\"/>"
    }

    /// 產生書籤結束標記
    func toBookmarkEndXML() -> String {
        return "<w:bookmarkEnd w:id=\"\(id)\"/>"
    }

    /// 產生完整的書籤 XML（包圍文字）
    /// 注意：這個方法產生的 XML 應該放在段落內、run 之間
    func toXML(wrapping text: String = "") -> String {
        if text.isEmpty {
            // 空書籤（只是一個標記點）
            return toBookmarkStartXML() + toBookmarkEndXML()
        } else {
            // 包圍文字的書籤
            return """
            \(toBookmarkStartXML())<w:r><w:t xml:space="preserve">\(escapeXML(text))</w:t></w:r>\(toBookmarkEndXML())
            """
        }
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

// MARK: - Bookmark Error

enum BookmarkError: Error, LocalizedError {
    case invalidName(String)
    case duplicateName(String)
    case notFound(String)

    var errorDescription: String? {
        switch self {
        case .invalidName(let name):
            return "Invalid bookmark name: '\(name)'. Names cannot contain spaces, start with numbers, or exceed 40 characters."
        case .duplicateName(let name):
            return "Bookmark with name '\(name)' already exists"
        case .notFound(let name):
            return "Bookmark '\(name)' not found"
        }
    }
}
