/* Neska native bridge: only runs inside the Android app (Capacitor).
   Android's WebView has no lock-screen media controls, so navigator.mediaSession
   is forwarded to the MediaSession plugin, which also keeps playback alive in the background. */
(function init(again) {
  var C = window.Capacitor;
  if (!C && !again) { document.addEventListener("DOMContentLoaded", function () { init(true); }); return; }
  if (!C || !C.isNativePlatform || !C.isNativePlatform() || C.getPlatform() !== "android") return;
  var P = C.Plugins && C.Plugins.MediaSession; if (!P) return;
  function q(fn) { try { var r = fn(); if (r && r.catch) r.catch(function () {}); } catch (e) {} }
  function art(a) { var o = []; (a || []).forEach(function (x) { try { if (x && x.src && !/^blob:|^data:/.test(x.src)) o.push({ src: new URL(x.src, location.href).href, sizes: x.sizes || "", type: x.type || "" }); } catch (e) {} }); return o; }
  var md = null, st = "none";
  var ms = {
    get metadata() { return md; },
    set metadata(m) { md = m; if (m) q(function () { return P.setMetadata({ title: m.title || "", artist: m.artist || "", album: m.album || "", artwork: art(m.artwork) }); }); },
    get playbackState() { return st; },
    set playbackState(s) { st = s || "none"; q(function () { return P.setPlaybackState({ playbackState: st }); }); },
    setActionHandler: function (a, h) { q(function () { return P.setActionHandler({ action: a }, h ? function (d) { try { h(Object.assign({ action: a }, d || {})); } catch (e) {} } : null); }); },
    setPositionState: function (p) { if (p && isFinite(p.duration)) q(function () { return P.setPositionState({ duration: p.duration, playbackRate: p.playbackRate || 1, position: p.position || 0 }); }); },
    setMicrophoneActive: function () {}, setCameraActive: function () {}
  };
  try { Object.defineProperty(navigator, "mediaSession", { value: ms, configurable: true }); } catch (e) {}
  if (typeof window.MediaMetadata === "undefined") window.MediaMetadata = function (o) { o = o || {}; this.title = o.title || ""; this.artist = o.artist || ""; this.album = o.album || ""; this.artwork = o.artwork || []; };
  document.documentElement.classList.add("is-native-android");
})(false);
