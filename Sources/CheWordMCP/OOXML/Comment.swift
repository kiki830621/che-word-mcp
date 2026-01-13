import Foundation

// MARK: - Comment

/// 註解（用於文件審閱和協作）
struct Comment {
    var id: Int                 // 註解唯一 ID
    var author: String          // 作者名稱
    var date: Date              // 建立日期
    var text: String            // 註解文字
    var initials: String?       // 作者縮寫（用於顯示）
    var paragraphIndex: Int     // 註解附加的段落索引

    init(id: Int, author: String, text: String, paragraphIndex: Int, date: Date = Date(), initials: String? = nil) {
        self.id = id
        self.author = author
        self.text = text
        self.paragraphIndex = paragraphIndex
        self.date = date
        self.initials = initials ?? String(author.prefix(2).uppercased())
    }
}

// MARK: - Comment XML Generation

extension Comment {
    /// 產生 comments.xml 中的單一註解 XML
    func toXML() -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: date)

        return """
        <w:comment w:id="\(id)" w:author="\(escapeXML(author))" w:date="\(dateString)" w:initials="\(escapeXML(initials ?? ""))">
            <w:p>
                <w:r>
                    <w:t xml:space="preserve">\(escapeXML(text))</w:t>
                </w:r>
            </w:p>
        </w:comment>
        """
    }

    /// 產生文件中的註解範圍開始標記
    func toCommentRangeStartXML() -> String {
        return "<w:commentRangeStart w:id=\"\(id)\"/>"
    }

    /// 產生文件中的註解範圍結束標記
    func toCommentRangeEndXML() -> String {
        return "<w:commentRangeEnd w:id=\"\(id)\"/>"
    }

    /// 產生文件中的註解參照標記
    func toCommentReferenceXML() -> String {
        return "<w:r><w:commentReference w:id=\"\(id)\"/></w:r>"
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

// MARK: - Comments Collection

/// 註解集合（用於管理文件中的所有註解）
struct CommentsCollection {
    var comments: [Comment] = []

    /// 取得下一個可用的註解 ID
    mutating func nextCommentId() -> Int {
        let maxId = comments.map { $0.id }.max() ?? 0
        return maxId + 1
    }

    /// 產生完整的 comments.xml 內容
    func toXML() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        """

        for comment in comments {
            xml += comment.toXML()
        }

        xml += "</w:comments>"
        return xml
    }

    /// Content Type for comments.xml
    static let contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

    /// Relationship type for comments
    static let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
}

// MARK: - Track Changes (Revision)

/// 修訂設定
struct TrackChangesSettings {
    var enabled: Bool = false           // 是否啟用修訂追蹤
    var author: String = "Unknown"      // 修訂作者
    var dateTime: Date = Date()         // 修訂時間

    init(enabled: Bool = false, author: String = "Unknown") {
        self.enabled = enabled
        self.author = author
        self.dateTime = Date()
    }
}

/// 修訂類型
enum RevisionType: String {
    case insertion = "ins"      // 插入
    case deletion = "del"       // 刪除
    case formatting = "rPrChange"   // 格式變更
    case paragraphChange = "pPrChange"  // 段落屬性變更
}

/// 單一修訂記錄
struct Revision {
    var id: Int                     // 修訂 ID
    var type: RevisionType          // 修訂類型
    var author: String              // 作者
    var date: Date                  // 修訂日期
    var paragraphIndex: Int         // 段落索引
    var originalText: String?       // 原始文字（刪除時）
    var newText: String?            // 新文字（插入時）

    init(id: Int, type: RevisionType, author: String, paragraphIndex: Int,
         originalText: String? = nil, newText: String? = nil, date: Date = Date()) {
        self.id = id
        self.type = type
        self.author = author
        self.paragraphIndex = paragraphIndex
        self.originalText = originalText
        self.newText = newText
        self.date = date
    }
}

// MARK: - Revision XML Generation

extension Revision {
    /// 產生插入標記的 XML
    func toInsertionXML(text: String) -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: date)

        return """
        <w:ins w:id="\(id)" w:author="\(escapeXML(author))" w:date="\(dateString)">
            <w:r>
                <w:t xml:space="preserve">\(escapeXML(text))</w:t>
            </w:r>
        </w:ins>
        """
    }

    /// 產生刪除標記的 XML
    func toDeletionXML(text: String) -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: date)

        return """
        <w:del w:id="\(id)" w:author="\(escapeXML(author))" w:date="\(dateString)">
            <w:r>
                <w:delText xml:space="preserve">\(escapeXML(text))</w:delText>
            </w:r>
        </w:del>
        """
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

// MARK: - Revisions Collection

/// 修訂集合
struct RevisionsCollection {
    var revisions: [Revision] = []
    var settings: TrackChangesSettings = TrackChangesSettings()

    /// 取得下一個修訂 ID
    mutating func nextRevisionId() -> Int {
        let maxId = revisions.map { $0.id }.max() ?? 0
        return maxId + 1
    }
}

// MARK: - Comment Error

enum CommentError: Error, LocalizedError {
    case notFound(Int)
    case invalidParagraphIndex(Int)

    var errorDescription: String? {
        switch self {
        case .notFound(let id):
            return "Comment with id \(id) not found"
        case .invalidParagraphIndex(let index):
            return "Invalid paragraph index: \(index)"
        }
    }
}

// MARK: - Revision Error

enum RevisionError: Error, LocalizedError {
    case notFound(Int)
    case trackChangesDisabled
    case cannotAccept(String)
    case cannotReject(String)

    var errorDescription: String? {
        switch self {
        case .notFound(let id):
            return "Revision with id \(id) not found"
        case .trackChangesDisabled:
            return "Track changes is not enabled"
        case .cannotAccept(let reason):
            return "Cannot accept revision: \(reason)"
        case .cannotReject(let reason):
            return "Cannot reject revision: \(reason)"
        }
    }
}
