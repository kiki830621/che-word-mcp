import Foundation
import ZIPFoundation

/// ZIP 壓縮/解壓縮工具
struct ZipHelper {
    /// 解壓縮 ZIP 檔案到臨時目錄
    static func unzip(_ url: URL) throws -> URL {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("che-word-mcp")
            .appendingPathComponent(UUID().uuidString)

        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        try FileManager.default.unzipItem(at: url, to: tempDir)

        return tempDir
    }

    /// 壓縮目錄內容為 ZIP 檔案（不包含目錄本身的路徑）
    static func zip(_ directory: URL, to destination: URL) throws {
        // 如果目標檔案已存在，先刪除
        if FileManager.default.fileExists(atPath: destination.path) {
            try FileManager.default.removeItem(at: destination)
        }

        // 建立新的 ZIP 檔案
        guard let archive = Archive(url: destination, accessMode: .create) else {
            throw WordError.zipError("無法建立 ZIP 檔案")
        }

        // 取得目錄內所有檔案的相對路徑
        let files = try getAllFiles(in: directory)

        for (relativePath, fileURL) in files {
            let fileData = try Data(contentsOf: fileURL)
            try archive.addEntry(
                with: relativePath,
                type: .file,
                uncompressedSize: Int64(fileData.count),
                compressionMethod: .deflate,
                provider: { position, size in
                    let startIndex = fileData.startIndex.advanced(by: Int(position))
                    let endIndex = startIndex.advanced(by: size)
                    return fileData.subdata(in: startIndex..<endIndex)
                }
            )
        }
    }

    /// 取得目錄內所有檔案（回傳相對路徑和完整 URL 的配對）
    private static func getAllFiles(in directory: URL) throws -> [(String, URL)] {
        var result: [(String, URL)] = []
        let fileManager = FileManager.default

        // 使用 subpathsOfDirectory 取得所有子路徑
        let directoryPath = directory.path
        let subpaths = try fileManager.subpathsOfDirectory(atPath: directoryPath)

        for subpath in subpaths {
            let fullURL = directory.appendingPathComponent(subpath)
            var isDirectory: ObjCBool = false
            if fileManager.fileExists(atPath: fullURL.path, isDirectory: &isDirectory) {
                if !isDirectory.boolValue {
                    result.append((subpath, fullURL))
                }
            }
        }

        return result
    }

    /// 清理臨時目錄
    static func cleanup(_ directory: URL) {
        try? FileManager.default.removeItem(at: directory)
    }
}
