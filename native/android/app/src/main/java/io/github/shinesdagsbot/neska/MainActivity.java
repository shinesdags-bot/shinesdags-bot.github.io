package io.github.shinesdagsbot.neska;

import android.os.Bundle;
import android.util.Log;
import androidx.webkit.WebViewCompat;
import androidx.webkit.WebViewFeature;
import com.getcapacitor.Bridge;
import com.getcapacitor.BridgeActivity;
import java.lang.reflect.Method;
import java.util.Collections;

public class MainActivity extends BridgeActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        injectBridgeIntoOfflinePage();
    }

    /**
     * Capacitor adds its JS bridge only to the live site's origin. The bundled offline screen
     * (errorPath, served from https://localhost) needs it too, to read the songs saved in the
     * app's storage and to drive the lock-screen controls. Best effort: on failure the offline
     * screen simply shows without songs, as before.
     */
    private void injectBridgeIntoOfflinePage() {
        try {
            Bridge b = this.bridge;
            if (b == null || !WebViewFeature.isFeatureSupported(WebViewFeature.DOCUMENT_START_SCRIPT)) return;
            String local = b.getScheme() + "://" + b.getHost();
            if (b.getServerUrl() == null || b.getServerUrl().startsWith(local)) return; // already covered
            Method getInjector = Bridge.class.getDeclaredMethod("getJSInjector");
            getInjector.setAccessible(true);
            Object injector = getInjector.invoke(b);
            if (injector == null) return;
            Method getScript = injector.getClass().getDeclaredMethod("getScriptString");
            getScript.setAccessible(true);
            String js = (String) getScript.invoke(injector);
            if (js == null || js.isEmpty()) return;
            WebViewCompat.addDocumentStartJavaScript(b.getWebView(), js, Collections.singleton(local));
        } catch (Throwable t) {
            Log.w("Neska", "Offline page bridge not added", t);
        }
    }
}
