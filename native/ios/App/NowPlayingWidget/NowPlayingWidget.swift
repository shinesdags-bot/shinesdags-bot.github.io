import WidgetKit
import SwiftUI
import AppIntents
import UIKit

// MARK: - Look

private enum Neska {
    static let plum = Color(red: 0x1a / 255, green: 0x0f / 255, blue: 0x1f / 255)
    static let plumHi = Color(red: 0x2e / 255, green: 0x18 / 255, blue: 0x36 / 255)
    static let rose = Color(red: 0xb7 / 255, green: 0x6e / 255, blue: 0x79 / 255)
    static let champagne = Color(red: 0xe8 / 255, green: 0xc4 / 255, blue: 0xa8 / 255)
    static let blush = Color(red: 0xc9 / 255, green: 0x8a / 255, blue: 0x9b / 255)
    static let roseGold = LinearGradient(colors: [champagne, blush, rose], startPoint: .topLeading, endPoint: .bottomTrailing)
    static let background = LinearGradient(colors: [plumHi, plum], startPoint: .top, endPoint: .bottom)
}

// MARK: - Timeline

struct NowEntry: TimelineEntry {
    let date: Date
    let state: NeskaNowState
    let cover: UIImage?
    let connected: Bool

    static let sample = NowEntry(date: Date(),
                                 state: NeskaNowState(title: "Golden Hour", artist: "Neska Music", playing: true),
                                 cover: nil, connected: true)
}

struct NowProvider: TimelineProvider {
    func placeholder(in context: Context) -> NowEntry { .sample }

    func getSnapshot(in context: Context, completion: @escaping (NowEntry) -> Void) {
        let e = current()
        completion(context.isPreview && !e.state.hasTrack ? .sample : e)
    }

    func getTimeline(in context: Context, completion: @escaping (Timeline<NowEntry>) -> Void) {
        // The app asks WidgetKit to reload whenever the song or play state changes.
        completion(Timeline(entries: [current()], policy: .never))
    }

    private func current() -> NowEntry {
        let s = NeskaShared.load()
        var img: UIImage?
        if s.hasTrack, let u = NeskaShared.coverURL { img = UIImage(contentsOfFile: u.path) }
        return NowEntry(date: Date(), state: s, cover: img, connected: NeskaShared.groupID != nil)
    }
}

// MARK: - Views

private struct CoverArt: View {
    let image: UIImage?
    let size: CGFloat
    let radius: CGFloat

    var body: some View {
        ZStack {
            if let image = image {
                Image(uiImage: image).resizable().aspectRatio(contentMode: .fill)
            } else {
                Rectangle().fill(Neska.roseGold)
                Image(systemName: "music.note")
                    .font(.system(size: size * 0.4, weight: .semibold))
                    .foregroundColor(Neska.plum.opacity(0.85))
            }
        }
        .frame(width: size, height: size)
        .clipShape(RoundedRectangle(cornerRadius: radius, style: .continuous))
        .overlay(RoundedRectangle(cornerRadius: radius, style: .continuous).stroke(Neska.champagne.opacity(0.18), lineWidth: 0.5))
        .shadow(color: .black.opacity(0.35), radius: 6, y: 3)
    }
}

private struct ControlButton<I: AppIntent>: View {
    let intent: I
    let symbol: String
    let size: CGFloat
    var primary = false
    let label: String

    var body: some View {
        Button(intent: intent) {
            ZStack {
                Circle().fill(primary ? AnyShapeStyle(Neska.roseGold) : AnyShapeStyle(Color.white.opacity(0.10)))
                Image(systemName: symbol)
                    .font(.system(size: size * (primary ? 0.42 : 0.38), weight: .bold))
                    .foregroundColor(primary ? Neska.plum : Neska.champagne)
                    .offset(x: symbol == "play.fill" ? size * 0.04 : 0)
            }
            .frame(width: size, height: size)
            .contentShape(Circle())
        }
        .buttonStyle(.plain)
        .accessibilityLabel(label)
    }
}

private struct TitleBlock: View {
    let entry: NowEntry
    let titleSize: CGFloat
    let titleLines: Int

    var body: some View {
        VStack(alignment: .leading, spacing: 2) {
            Text(entry.state.hasTrack ? entry.state.title : (entry.connected ? "Nothing playing" : "Open Neska"))
                .font(.system(size: titleSize, weight: .semibold, design: .rounded))
                .foregroundColor(.white)
                .lineLimit(titleLines)
                .minimumScaleFactor(0.85)
            Text(entry.state.hasTrack ? (entry.state.artist.isEmpty ? "Neska Music" : entry.state.artist)
                                      : (entry.connected ? "Tap play to start" : "to connect the widget"))
                .font(.system(size: titleSize * 0.8, weight: .medium, design: .rounded))
                .foregroundColor(Neska.champagne.opacity(0.85))
                .lineLimit(1)
        }
    }
}

struct SmallNowView: View {
    let entry: NowEntry
    var body: some View {
        VStack(alignment: .leading, spacing: 0) {
            HStack(alignment: .top, spacing: 0) {
                CoverArt(image: entry.cover, size: 66, radius: 14)
                Spacer(minLength: 6)
                VStack(spacing: 8) {
                    ControlButton(intent: NeskaTogglePlayIntent(), symbol: entry.state.playing ? "pause.fill" : "play.fill",
                                  size: 32, primary: true, label: entry.state.playing ? "Pause" : "Play")
                    ControlButton(intent: NeskaNextIntent(), symbol: "forward.fill", size: 28, label: "Next song")
                }
            }
            Spacer(minLength: 6)
            TitleBlock(entry: entry, titleSize: 14, titleLines: 1)
        }
        .padding(14)
    }
}

struct MediumNowView: View {
    let entry: NowEntry
    var body: some View {
        HStack(spacing: 14) {
            CoverArt(image: entry.cover, size: 126, radius: 18)
            VStack(alignment: .leading, spacing: 0) {
                HStack(spacing: 5) {
                    Image(systemName: entry.state.playing ? "waveform" : "pause")
                        .font(.system(size: 10, weight: .bold))
                    Text(entry.state.playing ? "NOW PLAYING" : "NESKA MUSIC")
                        .font(.system(size: 10, weight: .bold, design: .rounded))
                        .tracking(1.2)
                }
                .foregroundColor(Neska.blush)
                Spacer(minLength: 4)
                TitleBlock(entry: entry, titleSize: 17, titleLines: 2)
                Spacer(minLength: 6)
                HStack(spacing: 12) {
                    ControlButton(intent: NeskaPreviousIntent(), symbol: "backward.fill", size: 34, label: "Previous song")
                    ControlButton(intent: NeskaTogglePlayIntent(), symbol: entry.state.playing ? "pause.fill" : "play.fill",
                                  size: 42, primary: true, label: entry.state.playing ? "Pause" : "Play")
                    ControlButton(intent: NeskaNextIntent(), symbol: "forward.fill", size: 34, label: "Next song")
                }
            }
            Spacer(minLength: 0)
        }
        .padding(14)
    }
}

struct NowPlayingView: View {
    @Environment(\.widgetFamily) var family
    let entry: NowEntry
    var body: some View {
        Group {
            if family == .systemMedium { MediumNowView(entry: entry) } else { SmallNowView(entry: entry) }
        }
        .containerBackground(for: .widget) {
            ZStack {
                Neska.background
                RadialGradient(colors: [Neska.rose.opacity(0.28), .clear], center: .topTrailing, startRadius: 4, endRadius: 170)
            }
        }
    }
}

// MARK: - Widget

struct NowPlayingWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: NeskaShared.widgetKind, provider: NowProvider()) { entry in
            NowPlayingView(entry: entry)
        }
        .configurationDisplayName("Now Playing")
        .description("The song playing in Neska, with play, pause and skip.")
        .supportedFamilies([.systemSmall, .systemMedium])
        .contentMarginsDisabled()
    }
}

@main
struct NeskaWidgetBundle: WidgetBundle {
    var body: some Widget { NowPlayingWidget() }
}
