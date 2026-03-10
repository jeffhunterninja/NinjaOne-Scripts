//
// LoginView.swift
// NinjaOneITAM
//

import SwiftUI

struct LoginView: View {
    @ObservedObject var auth: NinjaOneAuth
    @State private var instance = "app.ninjarmm.com"
    @State private var clientId = ""
    @State private var isLoading = false
    @State private var errorMessage: String?

    var body: some View {
        Form {
            Section(header: Text("NinjaOne instance")) {
                TextField("Instance (e.g. app.ninjarmm.com)", text: $instance)
                    .textContentType(.URL)
                    .autocapitalization(.none)
                    .autocorrectionDisabled()
            }
            Section(header: Text("OAuth App (Native)")) {
                TextField("Client ID", text: $clientId)
                    .textContentType(.username)
                    .autocapitalization(.none)
                    .autocorrectionDisabled()
            }
            if let msg = errorMessage {
                Section {
                    Text(msg).foregroundColor(.red).font(.caption)
                }
            }
            Section {
                Button(action: doLogin) {
                    HStack {
                        if isLoading { ProgressView().scaleEffect(0.8) }
                        Text(isLoading ? "Signing in…" : "Log in with NinjaOne")
                    }
                    .frame(maxWidth: .infinity)
                }
                .disabled(isLoading || instance.isEmpty || clientId.isEmpty)
            }
        }
        .navigationTitle("NinjaOne ITAM")
    }

    private func doLogin() {
        errorMessage = nil
        isLoading = true
        Task {
            do {
                try await auth.startLogin(instance: instance, clientId: clientId)
            } catch {
                await MainActor.run {
                    errorMessage = error.localizedDescription
                    isLoading = false
                }
            }
            await MainActor.run { isLoading = false }
        }
    }
}
