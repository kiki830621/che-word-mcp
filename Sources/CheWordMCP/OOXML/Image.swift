import Foundation

// MARK: - Image Reference

/// 圖片參照（儲存在 word/media/ 目錄中的圖片）
struct ImageReference {
    var id: String           // 關係 ID (rId)
    var fileName: String     // 檔案名稱 (image1.png)
    var contentType: String  // MIME 類型 (image/png)
    var data: Data          // 圖片二進位資料

    init(id: String, fileName: String, contentType: String, data: Data) {
        self.id = id
        self.fileName = fileName
        self.contentType = contentType
        self.data = data
    }

    /// 從檔案路徑建立圖片參照
    static func from(path: String, id: String) throws -> ImageReference {
        let url = URL(fileURLWithPath: path)
        let data = try Data(contentsOf: url)

        let fileName = url.lastPathComponent
        let ext = url.pathExtension.lowercased()
        let contentType = mimeType(for: ext)

        return ImageReference(
            id: id,
            fileName: fileName,
            contentType: contentType,
            data: data
        )
    }

    /// 從 Base64 字串建立圖片參照
    static func from(base64: String, fileName: String, id: String) throws -> ImageReference {
        guard let data = Data(base64Encoded: base64) else {
            throw ImageError.invalidBase64
        }

        let ext = (fileName as NSString).pathExtension.lowercased()
        let contentType = mimeType(for: ext)

        return ImageReference(
            id: id,
            fileName: fileName,
            contentType: contentType,
            data: data
        )
    }

    /// 取得副檔名對應的 MIME 類型
    private static func mimeType(for ext: String) -> String {
        switch ext {
        case "png": return "image/png"
        case "jpg", "jpeg": return "image/jpeg"
        case "gif": return "image/gif"
        case "bmp": return "image/bmp"
        case "tiff", "tif": return "image/tiff"
        case "webp": return "image/webp"
        default: return "image/png"
        }
    }
}

// MARK: - Image Error

enum ImageError: Error, LocalizedError {
    case invalidBase64
    case fileNotFound(String)
    case unsupportedFormat(String)
    case dimensionRequired

    var errorDescription: String? {
        switch self {
        case .invalidBase64:
            return "Invalid Base64 encoded image data"
        case .fileNotFound(let path):
            return "Image file not found: \(path)"
        case .unsupportedFormat(let format):
            return "Unsupported image format: \(format)"
        case .dimensionRequired:
            return "Image dimensions (width/height) are required"
        }
    }
}

// MARK: - Drawing

/// 繪圖元素（用於將圖片嵌入文件）
struct Drawing {
    var type: DrawingType       // inline（行內）或 anchor（浮動）
    var width: Int              // 寬度（EMU）
    var height: Int             // 高度（EMU）
    var imageId: String         // 圖片關係 ID (rId)
    var name: String            // 圖片名稱
    var description: String     // 圖片描述（alt text）

    // 樣式屬性
    var hasBorder: Bool = false
    var borderColor: String = "000000"
    var borderWidth: Int = 9525  // EMU (約 0.75pt)
    var hasShadow: Bool = false

    init(type: DrawingType = .inline,
         width: Int,
         height: Int,
         imageId: String,
         name: String = "Picture",
         description: String = "") {
        self.type = type
        self.width = width
        self.height = height
        self.imageId = imageId
        self.name = name
        self.description = description
    }

    /// 從像素建立（1 像素 = 9525 EMU @ 96 DPI）
    static func from(widthPx: Int, heightPx: Int, imageId: String, name: String = "Picture") -> Drawing {
        return Drawing(
            width: widthPx * 9525,
            height: heightPx * 9525,
            imageId: imageId,
            name: name
        )
    }

    /// 從英寸建立（1 英寸 = 914400 EMU）
    static func from(widthInches: Double, heightInches: Double, imageId: String, name: String = "Picture") -> Drawing {
        return Drawing(
            width: Int(widthInches * 914400),
            height: Int(heightInches * 914400),
            imageId: imageId,
            name: name
        )
    }

    /// 從公分建立（1 公分 = 360000 EMU）
    static func from(widthCm: Double, heightCm: Double, imageId: String, name: String = "Picture") -> Drawing {
        return Drawing(
            width: Int(widthCm * 360000),
            height: Int(heightCm * 360000),
            imageId: imageId,
            name: name
        )
    }

    /// 取得寬度（像素）
    var widthInPixels: Int {
        width / 9525
    }

    /// 取得高度（像素）
    var heightInPixels: Int {
        height / 9525
    }
}

/// 繪圖類型
enum DrawingType {
    case inline    // 行內（隨文字流動）
    case anchor    // 浮動（絕對或相對定位）
}

// MARK: - XML Generation

extension Drawing {
    /// 轉換為 OOXML XML（完整的 drawing 元素，放在 Run 內）
    func toXML() -> String {
        switch type {
        case .inline:
            return toInlineXML()
        case .anchor:
            return toAnchorXML()
        }
    }

    /// 行內繪圖 XML
    private func toInlineXML() -> String {
        return """
        <w:drawing>
            <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="\(width)" cy="\(height)"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:docPr id="1" name="\(escapeXML(name))" descr="\(escapeXML(description))"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:nvPicPr>
                                <pic:cNvPr id="0" name="\(escapeXML(name))"/>
                                <pic:cNvPicPr/>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="\(imageId)" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="\(width)" cy="\(height)"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                                \(borderXML())
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:inline>
        </w:drawing>
        """
    }

    /// 浮動繪圖 XML（簡化版，使用 paragraph 錨點）
    private func toAnchorXML() -> String {
        return """
        <w:drawing>
            <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       distT="0" distB="0" distL="114300" distR="114300"
                       simplePos="0" relativeHeight="0" behindDoc="0" locked="0"
                       layoutInCell="1" allowOverlap="1">
                <wp:simplePos x="0" y="0"/>
                <wp:positionH relativeFrom="column">
                    <wp:posOffset>0</wp:posOffset>
                </wp:positionH>
                <wp:positionV relativeFrom="paragraph">
                    <wp:posOffset>0</wp:posOffset>
                </wp:positionV>
                <wp:extent cx="\(width)" cy="\(height)"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:wrapSquare wrapText="bothSides"/>
                <wp:docPr id="1" name="\(escapeXML(name))" descr="\(escapeXML(description))"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:nvPicPr>
                                <pic:cNvPr id="0" name="\(escapeXML(name))"/>
                                <pic:cNvPicPr/>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="\(imageId)" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="\(width)" cy="\(height)"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                                \(borderXML())
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:anchor>
        </w:drawing>
        """
    }

    /// 邊框 XML
    private func borderXML() -> String {
        guard hasBorder else { return "" }
        return """
        <a:ln w="\(borderWidth)">
            <a:solidFill>
                <a:srgbClr val="\(borderColor)"/>
            </a:solidFill>
        </a:ln>
        """
    }

    /// XML 跳脫
    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - Run with Drawing

extension Run {
    /// 建立含圖片的 Run
    static func withDrawing(_ drawing: Drawing) -> Run {
        var run = Run(text: "")
        run.drawing = drawing
        return run
    }
}
