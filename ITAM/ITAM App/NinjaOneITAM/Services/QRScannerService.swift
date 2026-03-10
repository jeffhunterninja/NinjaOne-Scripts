//
// QRScannerService.swift
// NinjaOneITAM
//
// Camera permission and QR metadata capture; parses JSON and returns ParsedQR.
//

import Foundation
import AVFoundation
import UIKit

/// Callback when a QR code string is read. Return true to keep scanning, false to stop.
typealias QRScanHandler = (String) -> Bool

final class QRScannerService: NSObject, ObservableObject {
    @Published var permissionGranted = false
    @Published var lastError: String?

    private var captureSession: AVCaptureSession?
    private var previewLayer: AVCaptureVideoPreviewLayer?
    private var onCode: QRScanHandler?

    func requestPermission() async -> Bool {
        let status = AVCaptureDevice.authorizationStatus(for: .video)
        switch status {
        case .authorized:
            await MainActor.run { permissionGranted = true }
            return true
        case .notDetermined:
            let granted = await AVCaptureDevice.requestAccess(for: .video)
            await MainActor.run { permissionGranted = granted }
            return granted
        case .denied, .restricted:
            await MainActor.run { lastError = "Camera access denied" }
            return false
        @unknown default:
            return false
        }
    }

    func buildPreviewLayer(in bounds: CGRect) -> AVCaptureVideoPreviewLayer? {
        let session = AVCaptureSession()
        guard let device = AVCaptureDevice.default(.builtInWideAngleCamera, for: .video, position: .back),
              let input = try? AVCaptureDeviceInput(device: device),
              session.canAddInput(input) else {
            return nil
        }
        session.addInput(input)
        let output = AVCaptureMetadataOutput()
        guard session.canAddOutput(output) else { return nil }
        session.addOutput(output)
        output.metadataObjectTypes = [.qr]
        output.setMetadataObjectsDelegate(self, queue: DispatchQueue.main)
        let layer = AVCaptureVideoPreviewLayer(session: session)
        layer.videoGravity = .resizeAspectFill
        layer.frame = bounds
        captureSession = session
        previewLayer = layer
        DispatchQueue.global(qos: .userInitiated).async { session.startRunning() }
        return layer
    }

    func setScanHandler(_ handler: @escaping QRScanHandler) {
        onCode = handler
    }

    func startSession() {
        captureSession?.startRunning()
    }

    func stopSession() {
        captureSession?.stopRunning()
    }

    func updatePreviewFrame(_ bounds: CGRect) {
        previewLayer?.frame = bounds
    }
}

extension QRScannerService: AVCaptureMetadataOutputObjectsDelegate {
    nonisolated func metadataOutput(_ output: AVCaptureMetadataOutput, didOutput metadataObjects: [AVMetadataObject], from connection: AVCaptureConnection) {
        guard let obj = metadataObjects.first as? AVMetadataMachineReadableCodeObject,
              let string = obj.stringValue else { return }
        DispatchQueue.main.async { [weak self] in
            let keepScanning = self?.onCode?(string) ?? true
            if !keepScanning {
                self?.captureSession?.stopRunning()
            }
        }
    }
}
