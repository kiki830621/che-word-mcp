import Foundation

// MARK: - Numbering

/// 編號定義集合
struct Numbering {
    var abstractNums: [AbstractNum] = []
    var nums: [Num] = []

    /// 取得下一個可用的 abstractNumId
    var nextAbstractNumId: Int {
        (abstractNums.map { $0.abstractNumId }.max() ?? -1) + 1
    }

    /// 取得下一個可用的 numId
    var nextNumId: Int {
        (nums.map { $0.numId }.max() ?? 0) + 1
    }

    init() {}

    /// 建立項目符號清單的編號定義
    mutating func createBulletList() -> Int {
        let abstractNumId = nextAbstractNumId
        let numId = nextNumId

        let abstractNum = AbstractNum(
            abstractNumId: abstractNumId,
            levels: [
                // Word 標準項目符號：Symbol 字體使用私有區字符
                Level(ilvl: 0, start: 1, numFmt: .bullet, lvlText: "\u{F0B7}", indent: 720, fontName: "Symbol"),  // ● 實心圓
                Level(ilvl: 1, start: 1, numFmt: .bullet, lvlText: "o", indent: 1440, fontName: "Courier New"),   // ○ 空心圓
                Level(ilvl: 2, start: 1, numFmt: .bullet, lvlText: "\u{F0A7}", indent: 2160, fontName: "Wingdings"), // ■ 實心方
                Level(ilvl: 3, start: 1, numFmt: .bullet, lvlText: "\u{F0B7}", indent: 2880, fontName: "Symbol"),
                Level(ilvl: 4, start: 1, numFmt: .bullet, lvlText: "o", indent: 3600, fontName: "Courier New"),
                Level(ilvl: 5, start: 1, numFmt: .bullet, lvlText: "\u{F0A7}", indent: 4320, fontName: "Wingdings"),
                Level(ilvl: 6, start: 1, numFmt: .bullet, lvlText: "\u{F0B7}", indent: 5040, fontName: "Symbol"),
                Level(ilvl: 7, start: 1, numFmt: .bullet, lvlText: "o", indent: 5760, fontName: "Courier New"),
                Level(ilvl: 8, start: 1, numFmt: .bullet, lvlText: "\u{F0A7}", indent: 6480, fontName: "Wingdings")
            ]
        )

        let num = Num(numId: numId, abstractNumId: abstractNumId)

        abstractNums.append(abstractNum)
        nums.append(num)

        return numId
    }

    /// 建立編號清單的編號定義
    mutating func createNumberedList() -> Int {
        let abstractNumId = nextAbstractNumId
        let numId = nextNumId

        let abstractNum = AbstractNum(
            abstractNumId: abstractNumId,
            levels: [
                Level(ilvl: 0, start: 1, numFmt: .decimal, lvlText: "%1.", indent: 720),
                Level(ilvl: 1, start: 1, numFmt: .lowerLetter, lvlText: "%2.", indent: 1440),
                Level(ilvl: 2, start: 1, numFmt: .lowerRoman, lvlText: "%3.", indent: 2160),
                Level(ilvl: 3, start: 1, numFmt: .decimal, lvlText: "%4.", indent: 2880),
                Level(ilvl: 4, start: 1, numFmt: .lowerLetter, lvlText: "%5.", indent: 3600),
                Level(ilvl: 5, start: 1, numFmt: .lowerRoman, lvlText: "%6.", indent: 4320),
                Level(ilvl: 6, start: 1, numFmt: .decimal, lvlText: "%7.", indent: 5040),
                Level(ilvl: 7, start: 1, numFmt: .lowerLetter, lvlText: "%8.", indent: 5760),
                Level(ilvl: 8, start: 1, numFmt: .lowerRoman, lvlText: "%9.", indent: 6480)
            ]
        )

        let num = Num(numId: numId, abstractNumId: abstractNumId)

        abstractNums.append(abstractNum)
        nums.append(num)

        return numId
    }
}

// MARK: - AbstractNum

/// 抽象編號定義（編號格式模板）
struct AbstractNum {
    var abstractNumId: Int
    var levels: [Level]

    init(abstractNumId: Int, levels: [Level] = []) {
        self.abstractNumId = abstractNumId
        self.levels = levels
    }
}

// MARK: - Num

/// 編號實例（引用 AbstractNum）
struct Num {
    var numId: Int
    var abstractNumId: Int

    init(numId: Int, abstractNumId: Int) {
        self.numId = numId
        self.abstractNumId = abstractNumId
    }
}

// MARK: - Level

/// 編號層級定義
struct Level {
    var ilvl: Int           // 層級（0-8）
    var start: Int          // 起始數字
    var numFmt: NumberFormat // 編號格式
    var lvlText: String     // 顯示格式（如 "%1.", "•"）
    var indent: Int         // 縮排（twips）
    var fontName: String?   // 字型（用於項目符號）

    init(ilvl: Int,
         start: Int = 1,
         numFmt: NumberFormat,
         lvlText: String,
         indent: Int,
         fontName: String? = nil) {
        self.ilvl = ilvl
        self.start = start
        self.numFmt = numFmt
        self.lvlText = lvlText
        self.indent = indent
        self.fontName = fontName
    }
}

// MARK: - NumberFormat

/// 編號格式
enum NumberFormat: String {
    case bullet = "bullet"              // 項目符號
    case decimal = "decimal"            // 1, 2, 3
    case lowerLetter = "lowerLetter"    // a, b, c
    case upperLetter = "upperLetter"    // A, B, C
    case lowerRoman = "lowerRoman"      // i, ii, iii
    case upperRoman = "upperRoman"      // I, II, III
    case none = "none"                  // 無編號
}

// MARK: - XML Generation

extension Numbering {
    /// 轉換為完整的 numbering.xml 內容
    func toXML() -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        """

        // 輸出所有 AbstractNum
        for abstractNum in abstractNums {
            xml += abstractNum.toXML()
        }

        // 輸出所有 Num
        for num in nums {
            xml += num.toXML()
        }

        xml += "</w:numbering>"
        return xml
    }
}

extension AbstractNum {
    func toXML() -> String {
        var xml = "<w:abstractNum w:abstractNumId=\"\(abstractNumId)\">"

        for level in levels {
            xml += level.toXML()
        }

        xml += "</w:abstractNum>"
        return xml
    }
}

extension Num {
    func toXML() -> String {
        return "<w:num w:numId=\"\(numId)\"><w:abstractNumId w:val=\"\(abstractNumId)\"/></w:num>"
    }
}

extension Level {
    func toXML() -> String {
        var xml = "<w:lvl w:ilvl=\"\(ilvl)\">"

        xml += "<w:start w:val=\"\(start)\"/>"
        xml += "<w:numFmt w:val=\"\(numFmt.rawValue)\"/>"
        xml += "<w:lvlText w:val=\"\(lvlText)\"/>"
        xml += "<w:lvlJc w:val=\"left\"/>"

        // 縮排設定
        xml += "<w:pPr>"
        xml += "<w:ind w:left=\"\(indent)\" w:hanging=\"360\"/>"
        xml += "</w:pPr>"

        // 項目符號的字型設定
        if numFmt == .bullet {
            let font = fontName ?? "Symbol"
            xml += "<w:rPr>"
            xml += "<w:rFonts w:ascii=\"\(font)\" w:hAnsi=\"\(font)\" w:hint=\"default\"/>"
            xml += "</w:rPr>"
        }

        xml += "</w:lvl>"
        return xml
    }
}
