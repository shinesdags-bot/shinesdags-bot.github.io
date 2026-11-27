import UIKit
import WebKit
import Capacitor

/// Bridge view controller that registers Neska's local plugins.
class NeskaViewController: CAPBridgeViewController {
    override open func capacitorDidLoad() {
        bridge?.registerPluginInstance(NeskaWidgetPlugin())
    }

    /// App-bound domains turn on the service worker (offline songs from the live site).
    /// Limit navigations only when the running bundle really lists WKAppBoundDomains:
    /// a host that drops the key (some sideload containers) would otherwise block every page.
    override open func webViewConfiguration(for instanceConfiguration: InstanceConfiguration) -> WKWebViewConfiguration {
        let config = super.webViewConfiguration(for: instanceConfiguration)
        if #available(iOS 14.0, *) {
            let domains = Bundle.main.object(forInfoDictionaryKey: "WKAppBoundDomains") as? [String] ?? []
            config.limitsNavigationsToAppBoundDomains = instanceConfiguration.limitsNavigationsToAppBoundDomains && !domains.isEmpty
        }
        return config
    }
}
