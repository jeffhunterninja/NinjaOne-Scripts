//
// QRScanView.swift
// NinjaOneITAM
//
// Full-screen camera preview with QR scanning; on scan, parses JSON and calls onUser or onDevice.
//

import SwiftUI
import AVFoundation
import UIKit

struct QRScanView: View {
    let mode: ScanMode
    let onUser: (UserQRPayload) -> Void
    let onDevice: (DeviceQRPayload) -> Void
    let onInvalid: (String) -> Void
    let onDismiss: () -> Void

    @StateObject private var scanner = QRScannerService()
    @State private var previewLayer: AVCaptureVideoPreviewLayer?

    enum ScanMode {
        case user
        case device
    }

    var body: some View {
        ZStack {
            CameraPreviewView(scanner: scanner)
                .ignoresSafeArea()
            VStack {
                Text(mode == .user ? "Scan user QR" : "Scan device QR")
                    .font(.headline)
                    .padding(8)
                    .background(.ultraThinMaterial)
                    .cornerRadius(8)
                Spacer()
            }
            .padding(.top, 40)
        }
        .onAppear {
            scanner.setScanHandler { [mode] string in
                let parsed = ParsedQR.parse(string)
                switch parsed {
                case .user(let payload):
                    if mode == .user {
                        onUser(payload)
                        return false
                    }
                    onInvalid("Expected device QR")
                    return true
                case .device(let payload):
                    if mode == .device {
                        onDevice(payload)
                        return false
                    }
                    onInvalid("Expected user QR")
                    return true
                case .invalid(let msg):
                    onInvalid(msg)
                    return true
                }
            }
            Task {
                let ok = await scanner.requestPermission()
                if !ok {
                    onInvalid("Camera access denied")
                    onDismiss()
                }
            }
        }
        .onDisappear { scanner.stopSession() }
        .overlay(alignment: .topLeading) {
            Button("Cancel") { onDismiss() }
                .padding()
                .foregroundColor(.white)
                .shadow(color: .black, radius: 2)
        }
    }
}

struct CameraPreviewView: UIViewRepresentable {
    @ObservedObject var scanner: QRScannerService

    func makeUIView(context: Context) -> UIView {
        let view = UIView(frame: .zero)
        view.backgroundColor = .black
        return view
    }

    func updateUIView(_ uiView: UIView, context: Context) {
        if scanner.permissionGranted && context.coordinator.layer == nil {
            let layer = scanner.buildPreviewLayer(in: uiView.bounds)
            if let layer = layer {
                layer.frame = uiView.bounds
                uiView.layer.addSublayer(layer)
                context.coordinator.layer = layer
            }
        }
        context.coordinator.layer?.frame = uiView.bounds
    }

    func makeCoordinator() -> Coordinator {
        Coordinator()
    }

    class Coordinator {
        var layer: AVCaptureVideoPreviewLayer?
    }
}
