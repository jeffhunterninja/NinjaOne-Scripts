//
// ContentView.swift
// NinjaOneITAM
//

import SwiftUI

struct ContentView: View {
    @ObservedObject var auth: NinjaOneAuth

    var body: some View {
        Group {
            if auth.isLoggedIn {
                AssignmentView(auth: auth)
            } else {
                LoginView(auth: auth)
            }
        }
    }
}
