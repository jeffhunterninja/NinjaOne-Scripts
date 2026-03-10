//
// AssignmentView.swift
// NinjaOneITAM
//
// Main screen: user + device list, scan user QR, scan device QR(s), submit.
//

import SwiftUI

struct AssignmentView: View {
    @ObservedObject var auth: NinjaOneAuth
    @State private var userPayload: UserQRPayload?
    @State private var devicePayloads: [DeviceQRPayload] = []
    @State private var showUserScanner = false
    @State private var showDeviceScanner = false
    @State private var isSubmitting = false
    @State private var submitResult: SubmitResult?
    @State private var toastMessage: String?
    @State private var showSettings = false

    private var api: NinjaOneAPI {
        NinjaOneAPI(instance: auth.instance, getToken: { try await auth.validAccessToken() })
    }

    var body: some View {
        NavigationStack {
            List {
                Section(header: Text("User")) {
                    if let u = userPayload {
                        HStack {
                            Text(u.displaySummary)
                            Spacer()
                            Button("Change") { showUserScanner = true }
                        }
                    } else {
                        Button("Scan user QR") { showUserScanner = true }
                    }
                }
                Section(header: Text("Devices to assign")) {
                    ForEach(Array(devicePayloads.enumerated()), id: \.offset) { _, d in
                        Text(d.displaySummary)
                    }
                    Button("Scan device QR") { showDeviceScanner = true }
                }
                if let r = submitResult {
                    Section(header: Text("Result")) {
                        Text(r.message).foregroundColor(r.success ? .primary : .red)
                    }
                }
            }
            .navigationTitle("Assign devices")
            .toolbar {
                ToolbarItem(placement: .topBarLeading) {
                    Button("Settings") { showSettings = true }
                }
                ToolbarItem(placement: .topBarTrailing) {
                    Button("Submit") { submit() }
                        .disabled(userPayload == nil || devicePayloads.isEmpty || isSubmitting)
                }
            }
            .fullScreenCover(isPresented: $showUserScanner) {
                QRScanView(
                    mode: .user,
                    onUser: { userPayload = $0; showUserScanner = false },
                    onDevice: { _ in toastMessage = "Expected user QR"; showUserScanner = false },
                    onInvalid: { toastMessage = $0 },
                    onDismiss: { showUserScanner = false }
                )
            }
            .fullScreenCover(isPresented: $showDeviceScanner) {
                QRScanView(
                    mode: .device,
                    onUser: { _ in toastMessage = "Expected device QR" },
                    onDevice: { devicePayloads.append($0); showDeviceScanner = false },
                    onInvalid: { toastMessage = $0 },
                    onDismiss: { showDeviceScanner = false }
                )
            }
            .sheet(isPresented: $showSettings) {
                SettingsView(auth: auth)
            }
            .overlay {
                if let msg = toastMessage {
                    VStack {
                        Spacer()
                        Text(msg)
                            .padding()
                            .background(.ultraThinMaterial)
                            .cornerRadius(8)
                            .padding()
                    }
                    .onTapGesture { toastMessage = nil }
                    .transition(.opacity)
                }
            }
        }
    }

    private func submit() {
        guard let user = userPayload else { return }
        isSubmitting = true
        submitResult = nil
        Task {
            let result = await runSubmit(user: user, devices: devicePayloads)
            await MainActor.run {
                submitResult = result
                isSubmitting = false
                if result.success {
                    userPayload = nil
                    devicePayloads = []
                }
            }
        }
    }
}

// MARK: - Submit flow

struct SubmitResult {
    let success: Bool
    let message: String
}

private func runSubmit(user: UserQRPayload, devices: [DeviceQRPayload]) async -> SubmitResult {
    let auth = NinjaOneAuth.shared
    let api = NinjaOneAPI(instance: auth.instance, getToken: { try await auth.validAccessToken() })

    do {
        let ownerUid: String
        if user.isEmailBased, let email = user.email {
            ownerUid = try await api.resolveOwnerUid(email: email)
        } else if let uid = user.uid {
            ownerUid = uid
        } else {
            return SubmitResult(success: false, message: "User has no email or UID")
        }

        var successCount = 0
        var errors: [String] = []
        for d in devices {
            do {
                _ = try await api.getDevice(id: d.id)
            } catch {
                errors.append("Device \(d.id): not found")
                continue
            }
            do {
                try await api.setDeviceOwner(deviceId: d.id, ownerUid: ownerUid)
                successCount += 1
            } catch {
                errors.append("Device \(d.id): \(error.localizedDescription)")
            }
        }

        if errors.isEmpty {
            return SubmitResult(success: true, message: "Assigned \(successCount) device(s) to user.")
        }
        let msg = "Assigned \(successCount); errors: \(errors.joined(separator: "; "))"
        return SubmitResult(success: successCount > 0, message: msg)
    } catch {
        return SubmitResult(success: false, message: error.localizedDescription)
    }
}

// MARK: - Settings

struct SettingsView: View {
    @ObservedObject var auth: NinjaOneAuth
    @Environment(\.dismiss) private var dismiss

    var body: some View {
        NavigationStack {
            Form {
                Section(header: Text("Instance")) {
                    Text(auth.instance.isEmpty ? "—" : auth.instance)
                }
                Section(header: Text("Client ID")) {
                    Text(auth.clientId.isEmpty ? "—" : auth.clientId)
                }
                Section {
                    Button("Log out", role: .destructive) {
                        auth.logout()
                        dismiss()
                    }
                }
            }
            .navigationTitle("Settings")
            .toolbar {
                ToolbarItem(placement: .confirmationAction) {
                    Button("Done") { dismiss() }
                }
            }
        }
    }
}
