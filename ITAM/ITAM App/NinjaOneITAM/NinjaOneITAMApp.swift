//
// NinjaOneITAMApp.swift
// NinjaOneITAM
//

import SwiftUI

@main
struct NinjaOneITAMApp: App {
    @StateObject private var auth = NinjaOneAuth.shared

    var body: some Scene {
        WindowGroup {
            ContentView(auth: auth)
        }
    }
}
