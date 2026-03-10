//
// NinjaOneAPI.swift
// NinjaOneITAM
//
// API client: users, contacts, device by id, set device owner.
//

import Foundation

// MARK: - API base

struct NinjaOneAPI {
    let baseURL: String
    let getToken: () async throws -> String

    init(instance: String, getToken: @escaping () async throws -> String) {
        let clean = instance.replacingOccurrences(of: "https://", with: "").replacingOccurrences(of: "/", with: "")
        baseURL = "https://\(clean)/ws/api/v2"
        self.getToken = getToken
    }

    private func request(path: String, method: String = "GET", body: Data? = nil) async throws -> (Data, HTTPURLResponse) {
        let token = try await getToken()
        let url = URL(string: "\(baseURL)/\(path)")!
        var req = URLRequest(url: url)
        req.httpMethod = method
        req.setValue("Bearer \(token)", forHTTPHeaderField: "Authorization")
        req.setValue("application/json", forHTTPHeaderField: "Accept")
        if let body = body {
            req.setValue("application/json", forHTTPHeaderField: "Content-Type")
            req.httpBody = body
        }
        let (data, response) = try await URLSession.shared.data(for: req)
        guard let http = response as? HTTPURLResponse else {
            throw NinjaOneAPIError.invalidResponse
        }
        return (data, http)
    }

    // GET /users
    func getUsers() async throws -> [NinjaOneUser] {
        let (data, http) = try await request(path: "users")
        guard (200...299).contains(http.statusCode) else {
            throw NinjaOneAPIError.httpStatus(http.statusCode, String(data: data, encoding: .utf8))
        }
        return try JSONDecoder().decode([NinjaOneUser].self, from: data)
    }

    // GET /contacts
    func getContacts() async throws -> [NinjaOneContact] {
        let (data, http) = try await request(path: "contacts")
        guard (200...299).contains(http.statusCode) else {
            throw NinjaOneAPIError.httpStatus(http.statusCode, String(data: data, encoding: .utf8))
        }
        return try JSONDecoder().decode([NinjaOneContact].self, from: data)
    }

    // GET /device/{id}
    func getDevice(id: Int) async throws -> NinjaOneDevice {
        let (data, http) = try await request(path: "device/\(id)")
        if http.statusCode == 404 {
            throw NinjaOneAPIError.deviceNotFound(id)
        }
        guard (200...299).contains(http.statusCode) else {
            throw NinjaOneAPIError.httpStatus(http.statusCode, String(data: data, encoding: .utf8))
        }
        return try JSONDecoder().decode(NinjaOneDevice.self, from: data)
    }

    // POST /device/{id}/owner/{ownerUid}
    func setDeviceOwner(deviceId: Int, ownerUid: String) async throws {
        let (data, http) = try await request(path: "device/\(deviceId)/owner/\(ownerUid)", method: "POST")
        if http.statusCode == 404 {
            throw NinjaOneAPIError.deviceNotFound(deviceId)
        }
        guard (200...299).contains(http.statusCode) else {
            throw NinjaOneAPIError.httpStatus(http.statusCode, String(data: data, encoding: .utf8))
        }
    }

    /// Resolve email to owner UID by checking users then contacts (matches Update-AssignedUser.ps1).
    func resolveOwnerUid(email: String) async throws -> String {
        let emailLower = email.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
        let users = try await getUsers()
        if let u = users.first(where: { ($0.email ?? "").trimmingCharacters(in: .whitespacesAndNewlines).lowercased() == emailLower }) {
            if let uid = u.uid { return "\(uid)" }
            if let id = u.id { return "\(id)" }
        }
        let contacts = try await getContacts()
        if let c = contacts.first(where: {
            let e = (c.email ?? c.Email ?? "").trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
            return e == emailLower
        }) {
            if let uid = c.uid { return "\(uid)" }
            if let id = c.id { return "\(id)" }
            if let Id = c.Id { return "\(Id)" }
        }
        throw NinjaOneAPIError.userNotFound(email)
    }
}

// MARK: - Errors

enum NinjaOneAPIError: LocalizedError {
    case invalidResponse
    case httpStatus(Int, String?)
    case deviceNotFound(Int)
    case userNotFound(String)

    var errorDescription: String? {
        switch self {
        case .invalidResponse: return "Invalid response"
        case .httpStatus(let code, let body): return "HTTP \(code): \(body ?? "")"
        case .deviceNotFound(let id): return "Device not found: \(id)"
        case .userNotFound(let email): return "User/contact not found: \(email)"
        }
    }
}

// MARK: - Models (API responses)

struct NinjaOneUser: Decodable {
    var id: Int?
    var uid: Int?
    var email: String?
}

struct NinjaOneContact: Decodable {
    var id: Int?
    var Id: Int?
    var uid: Int?
    var email: String?
    var Email: String?
}

struct NinjaOneDevice: Decodable {
    var id: Int?
    var systemName: String?
    var displayName: String?
}
