//
// NinjaOneAuth.swift
// NinjaOneITAM
//
// OAuth 2.0 Authorization Code + PKCE for NinjaOne. Keychain storage for tokens.
//

import Foundation
import AuthenticationServices
import Security
import CryptoKit
import UIKit

// MARK: - Keychain

enum KeychainKeys {
    static let service = "com.ninjaone.itam"
    static let accessToken = "accessToken"
    static let refreshToken = "refreshToken"
    static let tokenExpiry = "tokenExpiry"
    static let instance = "instance"
    static let clientId = "clientId"
}

final class KeychainStore {
    static let shared = KeychainStore()

    private func makeQuery(account: String, forWriting: Bool = false) -> [String: Any] {
        var query: [String: Any] = [
            kSecClass as String: kSecClassGenericPassword,
            kSecAttrService as String: KeychainKeys.service,
            kSecAttrAccount as String: account,
        ]
        if forWriting {
            query[kSecAttrAccessible as String] = kSecAttrAccessibleAfterFirstUnlockThisDeviceOnly
        }
        return query
    }

    func set(_ value: String, account: String) -> Bool {
        guard let data = value.data(using: .utf8) else { return false }
        var query = makeQuery(account: account, forWriting: true)
        query[kSecValueData as String] = data
        SecItemDelete(query as CFDictionary) // remove existing
        return SecItemAdd(query as CFDictionary, nil) == errSecSuccess
    }

    func set(_ value: Date, account: String) -> Bool {
        set(String(value.timeIntervalSince1970), account: account)
    }

    func set(_ value: Int, account: String) -> Bool {
        set(String(value), account: account)
    }

    func getString(account: String) -> String? {
        var query = makeQuery(account: account)
        query[kSecReturnData as String] = true
        query[kSecMatchLimit as String] = kSecMatchLimitOne
        var result: AnyObject?
        guard SecItemCopyMatching(query as CFDictionary, &result) == errSecSuccess,
              let data = result as? Data else { return nil }
        return String(data: data, encoding: .utf8)
    }

    func getDate(account: String) -> Date? {
        guard let s = getString(account: account), let t = TimeInterval(s) else { return nil }
        return Date(timeIntervalSince1970: t)
    }

    func remove(account: String) -> Bool {
        SecItemDelete(makeQuery(account: account) as CFDictionary) == errSecSuccess
    }

    func clearAll() {
        let accounts = [KeychainKeys.accessToken, KeychainKeys.refreshToken, KeychainKeys.tokenExpiry, KeychainKeys.instance, KeychainKeys.clientId]
        for account in accounts {
            remove(account: account)
        }
    }
}

// MARK: - PKCE

struct PKCE {
    let codeVerifier: String
    let codeChallenge: String

    static func generate() -> PKCE {
        var buffer = [UInt8](repeating: 0, count: 32)
        _ = SecRandomCopyBytes(kSecRandomDefault, buffer.count, &buffer)
        let verifier = Data(buffer).base64EncodedString()
            .replacingOccurrences(of: "+", with: "-")
            .replacingOccurrences(of: "/", with: "_")
            .replacingOccurrences(of: "=", with: "")
        let challenge = sha256Base64Url(verifier)
        return PKCE(codeVerifier: verifier, codeChallenge: challenge)
    }

    private static func sha256Base64Url(_ input: String) -> String {
        guard let data = input.data(using: .utf8) else { return "" }
        let hash = SHA256.hash(data: data)
        return Data(hash).base64EncodedString()
            .replacingOccurrences(of: "+", with: "-")
            .replacingOccurrences(of: "/", with: "_")
            .replacingOccurrences(of: "=", with: "")
    }
}

// MARK: - Token response

struct NinjaOneTokenResponse: Decodable {
    let accessToken: String
    let refreshToken: String?
    let expiresIn: Int?

    enum CodingKeys: String, CodingKey {
        case accessToken = "access_token"
        case refreshToken = "refresh_token"
        case expiresIn = "expires_in"
    }
}

// MARK: - Auth service

@MainActor
final class NinjaOneAuth: NSObject, ObservableObject {
    static let shared = NinjaOneAuth()

    @Published var isLoggedIn = false
    @Published var instance: String = ""
    @Published var clientId: String = ""

    private let keychain = KeychainStore.shared
    private var currentAccessToken: String?
    private var tokenExpiry: Date?

    override init() {
        super.init()
        loadStoredSession()
    }

    private func loadStoredSession() {
        instance = keychain.getString(account: KeychainKeys.instance) ?? ""
        clientId = keychain.getString(account: KeychainKeys.clientId) ?? ""
        currentAccessToken = keychain.getString(account: KeychainKeys.accessToken)
        tokenExpiry = keychain.getDate(account: KeychainKeys.tokenExpiry)
        isLoggedIn = currentAccessToken != nil && !instance.isEmpty
    }

    func saveConfig(instance: String, clientId: String) {
        self.instance = instance.trimmingCharacters(in: .whitespacesAndNewlines)
        self.clientId = clientId.trimmingCharacters(in: .whitespacesAndNewlines)
        _ = keychain.set(self.instance, account: KeychainKeys.instance)
        _ = keychain.set(self.clientId, account: KeychainKeys.clientId)
    }

    func startLogin(instance: String, clientId: String) async throws {
        saveConfig(instance: instance, clientId: clientId)
        let instanceClean = self.instance.replacingOccurrences(of: "https://", with: "").replacingOccurrences(of: "/", with: "")
        let redirectURI = "ninjaone-itam://callback"
        let scope = "monitoring management offline_access"
        let pkce = PKCE.generate()

        let authURL: URL = {
            var comp = URLComponents()
            comp.scheme = "https"
            comp.host = instanceClean
            comp.path = "/ws/oauth/authorize"
            comp.queryItems = [
                URLQueryItem(name: "response_type", value: "code"),
                URLQueryItem(name: "client_id", value: self.clientId),
                URLQueryItem(name: "redirect_uri", value: redirectURI),
                URLQueryItem(name: "scope", value: scope),
                URLQueryItem(name: "code_challenge", value: pkce.codeChallenge),
                URLQueryItem(name: "code_challenge_method", value: "S256"),
                URLQueryItem(name: "state", value: UUID().uuidString),
            ]
            return comp.url!
        }()

        let callbackURL = try await withCheckedThrowingContinuation { (cont: CheckedContinuation<URL, Error>) in
            let session = ASWebAuthenticationSession(
                url: authURL,
                callbackURLScheme: "ninjaone-itam"
            { callbackURL, error in
                if let error = error {
                    cont.resume(throwing: error)
                    return
                }
                guard let url = callbackURL else {
                    cont.resume(throwing: NSError(domain: "NinjaOneAuth", code: -1, userInfo: [NSLocalizedDescriptionKey: "No callback URL"]))
                    return
                }
                cont.resume(returning: url)
            }
            session.prefersEphemeralWebBrowserSession = false
            session.presentationContextProvider = self
            session.start()
        }

        guard let code = URLComponents(url: callbackURL, resolvingAgainstBaseURL: false)?.queryItems?.first(where: { $0.name == "code" })?.value else {
            throw NSError(domain: "NinjaOneAuth", code: -2, userInfo: [NSLocalizedDescriptionKey: "No authorization code in callback"])
        }

        try await exchangeCodeForTokens(code: code, redirectURI: redirectURI, codeVerifier: pkce.codeVerifier, instanceClean: instanceClean)
    }

    private func exchangeCodeForTokens(code: String, redirectURI: String, codeVerifier: String, instanceClean: String) async throws {
        let tokenURL = URL(string: "https://\(instanceClean)/ws/oauth/token")!
        var request = URLRequest(url: tokenURL)
        request.httpMethod = "POST"
        request.setValue("application/x-www-form-urlencoded", forHTTPHeaderField: "Content-Type")
        request.setValue("application/json", forHTTPHeaderField: "Accept")

        let body = [
            "grant_type": "authorization_code",
            "client_id": clientId,
            "code": code,
            "redirect_uri": redirectURI,
            "code_verifier": codeVerifier,
        ]
        request.httpBody = body
            .map { "\($0.key)=\($0.value.addingPercentEncoding(withAllowedCharacters: .urlQueryAllowed) ?? $0.value)" }
            .joined(separator: "&")
            .data(using: .utf8)

        let (data, response) = try await URLSession.shared.data(for: request)
        guard let http = response as? HTTPURLResponse, (200...299).contains(http.statusCode) else {
            let message = String(data: data, encoding: .utf8) ?? "Unknown error"
            throw NSError(domain: "NinjaOneAuth", code: -3, userInfo: [NSLocalizedDescriptionKey: "Token exchange failed: \(message)"])
        }

        let tokenResponse = try JSONDecoder().decode(NinjaOneTokenResponse.self, from: data)

        currentAccessToken = tokenResponse.accessToken
        let expiry = tokenResponse.expiresIn.map { Date().addingTimeInterval(TimeInterval($0)) }
        tokenExpiry = expiry

        _ = keychain.set(tokenResponse.accessToken, account: KeychainKeys.accessToken)
        if let refresh = tokenResponse.refreshToken {
            _ = keychain.set(refresh, account: KeychainKeys.refreshToken)
        }
        if let exp = expiry {
            _ = keychain.set(exp, account: KeychainKeys.tokenExpiry)
        }

        isLoggedIn = true
    }

    func refreshTokenIfNeeded() async throws {
        let refresh = keychain.getString(account: KeychainKeys.refreshToken)
        if let exp = tokenExpiry, Date() < exp.addingTimeInterval(-60) {
            return
        }
        guard let refresh = refresh else {
            if currentAccessToken == nil { throw NSError(domain: "NinjaOneAuth", code: -4, userInfo: [NSLocalizedDescriptionKey: "Not logged in"]) }
            return
        }

        let instanceClean = instance.replacingOccurrences(of: "https://", with: "").replacingOccurrences(of: "/", with: "")
        let tokenURL = URL(string: "https://\(instanceClean)/ws/oauth/token")!
        var request = URLRequest(url: tokenURL)
        request.httpMethod = "POST"
        request.setValue("application/x-www-form-urlencoded", forHTTPHeaderField: "Content-Type")
        request.setValue("application/json", forHTTPHeaderField: "Accept")

        let body = [
            "grant_type": "refresh_token",
            "client_id": clientId,
            "refresh_token": refresh,
        ]
        request.httpBody = body
            .map { "\($0.key)=\($0.value.addingPercentEncoding(withAllowedCharacters: .urlQueryAllowed) ?? $0.value)" }
            .joined(separator: "&")
            .data(using: .utf8)

        let (data, response) = try await URLSession.shared.data(for: request)
        guard let http = response as? HTTPURLResponse, (200...299).contains(http.statusCode) else {
            keychain.clearAll()
            currentAccessToken = nil
            tokenExpiry = nil
            isLoggedIn = false
            throw NSError(domain: "NinjaOneAuth", code: -5, userInfo: [NSLocalizedDescriptionKey: "Refresh failed; please log in again"])
        }

        let tokenResponse = try JSONDecoder().decode(NinjaOneTokenResponse.self, from: data)

        currentAccessToken = tokenResponse.accessToken
        let expiry = tokenResponse.expiresIn.map { Date().addingTimeInterval(TimeInterval($0)) }
        tokenExpiry = expiry
        _ = keychain.set(tokenResponse.accessToken, account: KeychainKeys.accessToken)
        if let exp = expiry {
            _ = keychain.set(exp, account: KeychainKeys.tokenExpiry)
        }
        if let newRefresh = tokenResponse.refreshToken {
            _ = keychain.set(newRefresh, account: KeychainKeys.refreshToken)
        }
    }

    func validAccessToken() async throws -> String {
        try await refreshTokenIfNeeded()
        guard let token = currentAccessToken else {
            throw NSError(domain: "NinjaOneAuth", code: -6, userInfo: [NSLocalizedDescriptionKey: "Not logged in"])
        }
        return token
    }

    func logout() {
        keychain.clearAll()
        currentAccessToken = nil
        tokenExpiry = nil
        isLoggedIn = false
    }
}

extension NinjaOneAuth: ASWebAuthenticationPresentationContextProviding {
    nonisolated func presentationAnchor(for session: ASWebAuthenticationSession) -> ASPresentationAnchor {
        let scenes = UIApplication.shared.connectedScenes
        guard let windowScene = scenes.compactMap({ $0 as? UIWindowScene }).first,
              let window = windowScene.windows.first(where: { $0.isKeyWindow }) ?? windowScene.windows.first else {
            return ASPresentationAnchor()
        }
        return window
    }
}
