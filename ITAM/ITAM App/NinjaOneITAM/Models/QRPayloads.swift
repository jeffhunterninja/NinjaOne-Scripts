//
// QRPayloads.swift
// NinjaOneITAM
//
// QR code JSON payloads: user (email or uid) and device (id, optional name).
//

import Foundation

enum QRPayloadType: String, Decodable {
    case user
    case device
}

/// User QR: {"type":"user","email":"..."} or {"type":"user","uid":"..."}
struct UserQRPayload: Decodable {
    let type: String
    let email: String?
    let uid: String?

    var displaySummary: String {
        if let e = email, !e.isEmpty { return e }
        if let u = uid, !u.isEmpty { return "UID: \(u)" }
        return "Unknown user"
    }

    var isEmailBased: Bool { email != nil && !(email?.isEmpty ?? true) }
}

/// Device QR: {"type":"device","id":12345} or {"type":"device","id":12345,"name":"DESKTOP-ABC"}
struct DeviceQRPayload: Decodable {
    let type: String
    let id: Int
    let name: String?

    var displaySummary: String {
        if let n = name, !n.isEmpty { return "\(n) (ID: \(id))" }
        return "Device \(id)"
    }
}

/// Parsed result from a single QR scan.
enum ParsedQR {
    case user(UserQRPayload)
    case device(DeviceQRPayload)
    case invalid(String)

    static func parse(_ string: String) -> ParsedQR {
        let trimmed = string.trimmingCharacters(in: .whitespacesAndNewlines)
        guard let data = trimmed.data(using: .utf8) else { return .invalid("Invalid encoding") }
        do {
            if let raw = try JSONSerialization.jsonObject(with: data) as? [String: Any],
               let type = raw["type"] as? String {
                if type == "user" {
                    let payload = try JSONDecoder().decode(UserQRPayload.self, from: data)
                    return .user(payload)
                }
                if type == "device" {
                    let payload = try JSONDecoder().decode(DeviceQRPayload.self, from: data)
                    return .device(payload)
                }
            }
            return .invalid("Unknown QR type")
        } catch {
            return .invalid(error.localizedDescription)
        }
    }
}
