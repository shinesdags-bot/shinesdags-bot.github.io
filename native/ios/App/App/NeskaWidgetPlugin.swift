import Foundation
import UIKit
import Capacitor

/// Feeds the Now Playing home-screen widget and relays its buttons back to the web player.
/// JS: Capacitor.Plugins.NeskaWidget.update({title, artist, playing, artKey, artUrls:[...], artData:"<base64 jpeg>"})
///     Capacitor.Plugins.NeskaWidget.addListener("widgetAction", ({action}) => ...)  // toggle | next | prev
@objc(NeskaWidgetPlugin)
public class NeskaWidgetPlugin: CAPPlugin, CAPBridgedPlugin {
    public let identifier = "NeskaWidgetPlugin"
    public let jsName = "NeskaWidget"
    public let pluginMethods: [CAPPluginMethod] = [
        CAPPluginMethod(name: "update", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "info", returnType: CAPPluginReturnPromise)
    ]

    private static weak var shared: NeskaWidgetPlugin?
    private var state = NeskaNowState()
    private var artTask: URLSessionDataTask?
    private var pendingReload: DispatchWorkItem?

    override public func load() {
        NeskaWidgetPlugin.shared = self
        state = NeskaShared.load()
        NotificationCenter.default.addObserver(self, selector: #selector(onAction(_:)), name: NeskaShared.actionNote, object: nil)
        let center = CFNotificationCenterGetDarwinNotifyCenter()
        for a in NeskaShared.actions {
            CFNotificationCenterAddObserver(center, nil, { _, _, name, _, _ in
                guard let cf = name?.rawValue else { return }
                let n = cf as String
                let action = String(n.dropFirst(NeskaShared.darwinPrefix.count))
                DispatchQueue.main.async { NeskaWidgetPlugin.shared?.relay(action) }
            }, (NeskaShared.darwinPrefix + a) as CFString, nil, .deliverImmediately)
        }
    }

    @objc private func onAction(_ n: Notification) {
        guard let a = n.userInfo?["action"] as? String else { return }
        relay(a)
    }

    private func relay(_ action: String) {
        if action == "toggle" { state = NeskaShared.load() }   // the intent flipped it optimistically
        notifyListeners("widgetAction", data: ["action": action], retainUntilConsumed: true)
    }

    @objc func info(_ call: CAPPluginCall) {
        call.resolve(["group": NeskaShared.groupID ?? "", "available": NeskaShared.groupID != nil])
    }

    @objc func update(_ call: CAPPluginCall) {
        guard NeskaShared.groupID != nil else { call.resolve(["ok": false, "reason": "no app group"]); return }
        var s = state
        if let t = call.getString("title") { s.title = t }
        if let a = call.getString("artist") { s.artist = a }
        if let p = call.getBool("playing") { s.playing = p }
        let key = call.getString("artKey")
        let artData = call.getString("artData")
        let urls = (call.getArray("artUrls", String.self) ?? []).filter { $0.hasPrefix("http") }
        call.resolve(["ok": true])

        if let key = key, key != s.artKey {
            s.artKey = key
            artTask?.cancel(); artTask = nil
            if let b64 = artData, !b64.isEmpty {
                let raw = b64.components(separatedBy: ",").last ?? b64
                writeCover(Data(base64Encoded: raw).flatMap { UIImage(data: $0) })
            } else if !urls.isEmpty {
                fetchCover(urls, key: key)
            } else {
                writeCover(nil)
            }
        }
        if s != state || s.updated == 0 {
            s.updated = Date().timeIntervalSince1970
            state = s
            NeskaShared.save(s)
            scheduleReload()
        }
    }

    private func fetchCover(_ urls: [String], key: String) {
        guard let first = urls.first, let u = URL(string: first) else { writeCover(nil); return }
        let rest = Array(urls.dropFirst())
        artTask = URLSession.shared.dataTask(with: u) { [weak self] data, resp, err in
            guard let self = self else { return }
            let ok = (resp as? HTTPURLResponse).map { (200..<300).contains($0.statusCode) } ?? false
            let img = ok ? data.flatMap { UIImage(data: $0) } : nil
            DispatchQueue.main.async {
                guard self.state.artKey == key else { return }
                if let img = img, img.size.width > 130 { self.writeCover(img) }   // YouTube's 404 stub is 120x90
                else if !rest.isEmpty { self.fetchCover(rest, key: key) }
                else { self.writeCover(img) }
            }
        }
        artTask?.resume()
    }

    private func writeCover(_ img: UIImage?) {
        guard let url = NeskaShared.coverURL else { return }
        if let img = img, let jpg = NeskaWidgetPlugin.square(img, side: 300).jpegData(compressionQuality: 0.85) {
            try? jpg.write(to: url, options: .atomic)
        } else {
            try? FileManager.default.removeItem(at: url)
        }
        scheduleReload()
    }

    /// Center-crops to a square and scales to `side` px (keeps the widget well under its memory limit).
    static func square(_ img: UIImage, side: CGFloat) -> UIImage {
        let w = img.size.width, h = img.size.height
        guard w > 0, h > 0 else { return img }
        let scale = side / min(w, h)
        let fmt = UIGraphicsImageRendererFormat.default()
        fmt.scale = 1
        fmt.opaque = true
        return UIGraphicsImageRenderer(size: CGSize(width: side, height: side), format: fmt).image { _ in
            let dw = w * scale, dh = h * scale
            img.draw(in: CGRect(x: (side - dw) / 2, y: (side - dh) / 2, width: dw, height: dh))
        }
    }

    private func scheduleReload() {
        pendingReload?.cancel()
        let w = DispatchWorkItem { NeskaShared.reloadWidget() }
        pendingReload = w
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.35, execute: w)
    }
}
