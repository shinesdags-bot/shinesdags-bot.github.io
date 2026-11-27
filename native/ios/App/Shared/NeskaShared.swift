// Shared by the Neska app and its Now Playing widget (compiled into both targets).
import Foundation
#if canImport(WidgetKit)
import WidgetKit
#endif
#if canImport(AppIntents)
import AppIntents
#endif

struct NeskaNowState: Codable, Equatable {
    var title: String = ""
    var artist: String = ""
    var playing: Bool = false
    var artKey: String = ""      // identifies the cover currently saved as cover.jpg
    var updated: Double = 0
    var hasTrack: Bool { !title.isEmpty }
}

enum NeskaShared {
    static let baseGroup = "group.io.github.shinesdagsbot.neska"
    static let widgetKind = "NeskaNowPlaying"
    static let actionNote = Notification.Name("NeskaWidgetAction")
    static let darwinPrefix = "io.github.shinesdagsbot.neska.widget."
    static let actions = ["toggle", "next", "prev"]

    /// True inside the widget extension process.
    static var isExtension: Bool { Bundle.main.bundleURL.pathExtension == "appex" }

    /// The app's own bundle (the containing app when running in the extension).
    static var appBundle: Bundle {
        if isExtension {
            let appURL = Bundle.main.bundleURL.deletingLastPathComponent().deletingLastPathComponent()
            if let b = Bundle(url: appURL) { return b }
        }
        return Bundle.main
    }

    /// App group actually granted to this install. SideStore/AltStore re-sign with a
    /// team-specific group (group.<id>.<TEAMID>) and list it in Info.plist under ALTAppGroups.
    static let groupID: String? = {
        var c: [String] = []
        func alt(_ b: Bundle) -> [String] {
            let v = b.object(forInfoDictionaryKey: "ALTAppGroups")
            if let a = v as? [String] { return a }
            if let s = v as? String { return [s] }
            return []
        }
        let groups = alt(Bundle.main) + alt(appBundle)
        c += groups.filter { $0.contains("shinesdagsbot.neska") }
        c += groups
        if let id = appBundle.bundleIdentifier { c.append("group." + id) }
        c.append(baseGroup)
        var seen = Set<String>()
        for g in c where !g.isEmpty && seen.insert(g).inserted {
            if FileManager.default.containerURL(forSecurityApplicationGroupIdentifier: g) != nil { return g }
        }
        return nil
    }()

    static var container: URL? {
        guard let g = groupID else { return nil }
        return FileManager.default.containerURL(forSecurityApplicationGroupIdentifier: g)
    }
    static var stateURL: URL? { container?.appendingPathComponent("neska-now.json") }
    static var coverURL: URL? { container?.appendingPathComponent("neska-cover.jpg") }

    static func load() -> NeskaNowState {
        guard let u = stateURL, let d = try? Data(contentsOf: u),
              let s = try? JSONDecoder().decode(NeskaNowState.self, from: d) else { return NeskaNowState() }
        return s
    }

    static func save(_ s: NeskaNowState) {
        guard let u = stateURL, let d = try? JSONEncoder().encode(s) else { return }
        try? d.write(to: u, options: .atomic)
    }

    static func reloadWidget() {
        #if canImport(WidgetKit)
        if #available(iOS 14.0, *) { WidgetCenter.shared.reloadTimelines(ofKind: widgetKind) }
        #endif
    }

    /// Route a widget button press to the player inside the app.
    static func dispatch(_ action: String) {
        if action == "toggle" {
            var s = load(); s.playing.toggle(); s.updated = Date().timeIntervalSince1970; save(s)
        }
        if isExtension {
            // Running in the widget process: wake the app's observer over Darwin notify.
            let name = CFNotificationName((darwinPrefix + action) as CFString)
            CFNotificationCenterPostNotification(CFNotificationCenterGetDarwinNotifyCenter(), name, nil, nil, true)
        } else {
            DispatchQueue.main.async {
                NotificationCenter.default.post(name: actionNote, object: nil, userInfo: ["action": action])
            }
        }
    }
}

#if canImport(AppIntents)
@available(iOS 17.0, *)
struct NeskaTogglePlayIntent: AudioPlaybackIntent {
    static let title: LocalizedStringResource = "Play or Pause Neska"
    static let isDiscoverable: Bool = false
    init() {}
    func perform() async throws -> some IntentResult {
        NeskaShared.dispatch("toggle")
        return .result()
    }
}

@available(iOS 17.0, *)
struct NeskaNextIntent: AudioPlaybackIntent {
    static let title: LocalizedStringResource = "Next Song"
    static let isDiscoverable: Bool = false
    init() {}
    func perform() async throws -> some IntentResult {
        NeskaShared.dispatch("next")
        return .result()
    }
}

@available(iOS 17.0, *)
struct NeskaPreviousIntent: AudioPlaybackIntent {
    static let title: LocalizedStringResource = "Previous Song"
    static let isDiscoverable: Bool = false
    init() {}
    func perform() async throws -> some IntentResult {
        NeskaShared.dispatch("prev")
        return .result()
    }
}
#endif
