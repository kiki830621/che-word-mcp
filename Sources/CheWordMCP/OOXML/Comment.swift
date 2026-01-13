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

    // 回覆支援（Word 2012+ / commentsExtended.xml）
    var paraId: String?         // 段落 ID（用於連結回覆）
    var parentId: Int?          // 父註解 ID（如果是回覆）
    var done: Bool = false      // 是否已解決

    init(id: Int, author: String, text: String, paragraphIndex: Int, date: Date = Date(), initials: String? = nil) {
        self.id = id
        self.author = author
        self.text = text
        self.paragraphIndex = paragraphIndex
        self.date = date
        self.initials = initials ?? String(author.prefix(2).uppercased())
        self.paraId = Comment.generateParaId()
    }

    /// 建立回覆註解
    init(id: Int, author: String, text: String, parentId: Int, date: Date = Date(), initials: String? = nil) {
        self.id = id
        self.author = author
        self.text = text
        self.paragraphIndex = -1  // 回覆不直接附加到段落
        self.parentId = parentId
        self.date = date
        self.initials = initials ?? String(author.prefix(2).uppercased())
        self.paraId = Comment.generateParaId()
    }

    /// 產生隨機 8 位十六進位段落 ID
    private static func generateParaId() -> String {
        return String(format: "%08X", UInt32.random(in: 0...UInt32.max))
    }

    /// 是否為回覆
    var isReply: Bool {
        return parentId != nil
    }
}

// MARK: - Comment XML Generation

extension Comment {
    /// 產生 comments.xml 中的單一註解 XML
    func toXML() -> String {
        let dateFormatter = ISO8601DateFormatter()
        let dateString = dateFormatter.string(from: date)

        // 段落屬性（包含 paraId）
        var pPrAttrs = ""
        if let paraId = paraId {
            pPrAttrs = " w14:paraId=\"\(paraId)\" w14:textId=\"\(String(format: "%08X", UInt32.random(in: 0...UInt32.max)))\""
        }

        return """
        <w:comment w:id="\(id)" w:author="\(escapeXML(author))" w:date="\(dateString)" w:initials="\(escapeXML(initials ?? ""))">
            <w:p\(pPrAttrs)>
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

    /// 新增註解
    mutating func addComment(_ comment: Comment) {
        comments.append(comment)
    }

    /// 新增回覆
    mutating func addReply(to parentId: Int, author: String, text: String) -> Comment? {
        guard comments.contains(where: { $0.id == parentId }) else {
            return nil
        }
        let newId = nextCommentId()
        let reply = Comment(id: newId, author: author, text: text, parentId: parentId)
        comments.append(reply)
        return reply
    }

    /// 取得某註解的所有回覆
    func getReplies(for commentId: Int) -> [Comment] {
        return comments.filter { $0.parentId == commentId }
    }

    /// 取得頂層註解（非回覆）
    func getTopLevelComments() -> [Comment] {
        return comments.filter { !$0.isReply }
    }

    /// 標記註解為已解決
    mutating func markAsDone(_ commentId: Int, done: Bool = true) {
        if let index = comments.firstIndex(where: { $0.id == commentId }) {
            comments[index].done = done
        }
    }

    /// 產生完整的 comments.xml 內容
    func toXML() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        """

        for comment in comments {
            xml += comment.toXML()
        }

        xml += "</w:comments>"
        return xml
    }

    /// 產生 commentsExtended.xml 內容（Word 2012+ 回覆支援）
    func toExtendedXML() -> String? {
        // 只有當有回覆或已解決狀態時才需要 commentsExtended.xml
        let hasExtendedInfo = comments.contains { $0.isReply || $0.done }
        guard hasExtendedInfo else { return nil }

        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                        mc:Ignorable="w15">
        """

        for comment in comments {
            var attrs: [String] = ["w15:paraId=\"\(comment.paraId ?? "00000000")\""]

            // 如果是回覆，找到父註解的 paraId
            if let parentId = comment.parentId,
               let parentComment = comments.first(where: { $0.id == parentId }),
               let parentParaId = parentComment.paraId {
                attrs.append("w15:paraIdParent=\"\(parentParaId)\"")
            }

            // 如果已解決
            if comment.done {
                attrs.append("w15:done=\"1\"")
            }

            xml += "<w15:commentEx \(attrs.joined(separator: " "))/>"
        }

        xml += "</w15:commentsEx>"
        return xml
    }

    /// 是否有需要 commentsExtended.xml
    var hasExtendedComments: Bool {
        return comments.contains { $0.isReply || $0.done }
    }

    /// Content Type for comments.xml
    static let contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

    /// Content Type for commentsExtended.xml
    static let extendedContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"

    /// Relationship type for comments
    static let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"

    /// Relationship type for commentsExtended
    static let extendedRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
}

// MARK: - Comment Error

enum CommentError: Error, LocalizedError {
    case notFound(Int)
    case invalidParagraphIndex(Int)
    case parentCommentNotFound(Int)
    case cannotReplyToReply  // 視實作需求可移除此限制

    var errorDescription: String? {
        switch self {
        case .notFound(let id):
            return "Comment with id \(id) not found"
        case .invalidParagraphIndex(let index):
            return "Invalid paragraph index: \(index)"
        case .parentCommentNotFound(let id):
            return "Parent comment with id \(id) not found"
        case .cannotReplyToReply:
            return "Cannot reply to a reply (nested replies not supported)"
        }
    }
}

