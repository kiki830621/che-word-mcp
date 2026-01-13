import Foundation
import MCP

// Entry point
let server = WordMCPServer()
try await server.run()
