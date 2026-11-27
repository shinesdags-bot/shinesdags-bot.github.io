import UIKit
import Capacitor

/// Bridge view controller that registers Neska's local plugins.
class NeskaViewController: CAPBridgeViewController {
    override open func capacitorDidLoad() {
        bridge?.registerPluginInstance(NeskaWidgetPlugin())
    }
}
