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

/* Neska iOS widget feed: only runs inside the iOS app (Capacitor).
   Mirrors navigator.mediaSession (song, artist, cover, play state) to the Now Playing
   home-screen widget via the NeskaWidget plugin, and runs the widget's buttons through
   the player's own media-session handlers. WebKit's native lock-screen session is kept. */
(function initIOS(again) {
  var C = window.Capacitor;
  if (!C && !again) { document.addEventListener("DOMContentLoaded", function () { initIOS(true); }); return; }
  if (!C || !C.isNativePlatform || !C.isNativePlatform() || C.getPlatform() !== "ios") return;
  var W = C.Plugins && C.Plugins.NeskaWidget; if (!W || !W.update) return;
  var handlers = {}, md = null, mdSet = false, st = "none", last = "", artSent = null, timer = 0;
  function q(fn) { try { var r = fn(); if (r && r.catch) r.catch(function () {}); } catch (e) {} }
  function abs(u) { try { return new URL(u, location.href).href; } catch (e) { return ""; } }
  function srcs(m) { var o = []; try { [].forEach.call((m && m.artwork) || [], function (x) { if (x && x.src) { var u = /^(blob|data):/.test(x.src) ? x.src : abs(x.src); if (u) o.push(u); } }); } catch (e) {} return o; }
  function toJpeg(src, cb) {
    var im = new Image(), done = false;
    function fin(v) { if (!done) { done = true; cb(v); } }
    im.onload = function () {
      try {
        var s = 300, c = document.createElement("canvas"), g = c.getContext("2d"), w = im.naturalWidth, h = im.naturalHeight, k = s / Math.min(w, h);
        c.width = c.height = s; g.fillStyle = "#1a0f1f"; g.fillRect(0, 0, s, s);
        g.drawImage(im, (s - w * k) / 2, (s - h * k) / 2, w * k, h * k);
        fin(c.toDataURL("image/jpeg", 0.85));
      } catch (e) { fin(""); }
    };
    im.onerror = function () { fin(""); };
    setTimeout(function () { fin(""); }, 8000);
    im.src = src;
  }
  function send() {
    timer = 0;
    var p = { playing: st === "playing" };
    if (mdSet) {
      p.title = (md && md.title) || ""; p.artist = (md && md.artist) || "";
      var list = srcs(md), key = list[0] || "";
      p.artKey = key;
      if (key !== artSent) {
        var local = /^(blob|data):/.test(key);
        if (local) {
          artSent = key;
          toJpeg(key, function (d) { if (artSent !== key) return; q(function () { return W.update({ title: p.title, artist: p.artist, playing: st === "playing", artKey: key, artData: d }); }); });
          delete p.artKey; // sent with the converted image
        } else { artSent = key; p.artUrls = list.filter(function (u) { return /^https?:/.test(u); }); }
      }
    }
    var sig = JSON.stringify(p);
    if (sig === last) return;
    last = sig;
    q(function () { return W.update(p); });
  }
  function push() { if (!timer) timer = setTimeout(send, 250); }

  var MS = window.MediaSession, P = MS && MS.prototype;
  var dm = P && Object.getOwnPropertyDescriptor(P, "metadata"), dp = P && Object.getOwnPropertyDescriptor(P, "playbackState");
  if (navigator.mediaSession && dm && dm.set && dp && dp.set && P.setActionHandler) {
    var sah = P.setActionHandler;
    Object.defineProperty(P, "metadata", { configurable: true, enumerable: dm.enumerable, get: dm.get,
      set: function (v) { try { dm.set.call(this, v); } finally { md = v; mdSet = true; push(); } } });
    Object.defineProperty(P, "playbackState", { configurable: true, enumerable: dp.enumerable, get: dp.get,
      set: function (v) { try { dp.set.call(this, v); } finally { st = v || "none"; push(); } } });
    P.setActionHandler = function (a, h) { handlers[a] = h || null; return sah.apply(this, arguments); };
  } else {
    // No WebKit media session: a stand-in so the player's calls still reach the widget.
    var ms = {
      get metadata() { return md; }, set metadata(v) { md = v; mdSet = true; push(); },
      get playbackState() { return st; }, set playbackState(v) { st = v || "none"; push(); },
      setActionHandler: function (a, h) { handlers[a] = h || null; },
      setPositionState: function () {}, setMicrophoneActive: function () {}, setCameraActive: function () {}
    };
    try { Object.defineProperty(navigator, "mediaSession", { value: ms, configurable: true }); } catch (e) {}
    if (typeof window.MediaMetadata === "undefined") window.MediaMetadata = function (o) { o = o || {}; this.title = o.title || ""; this.artist = o.artist || ""; this.album = o.album || ""; this.artwork = o.artwork || []; };
  }
  ["play", "pause", "ended"].forEach(function (ev) { document.addEventListener(ev, push, true); });

  function run(name) {
    var h = handlers[name];
    if (h) { try { h({ action: name }); return; } catch (e) {} }
    var NP = window.NeskaPlayer, a = NP && NP.audio && NP.audio();
    if (a && (name === "play" || name === "pause")) { try { if (name === "play") { var r = a.play(); if (r && r.catch) r.catch(function () {}); } else a.pause(); } catch (e) {} }
  }
  q(function () {
    return W.addListener("widgetAction", function (e) {
      var a = e && e.action;
      if (a === "toggle") run(st === "playing" ? "pause" : "play");
      else if (a === "next") run("nexttrack");
      else if (a === "prev") run("previoustrack");
      // re-send the real state so the widget settles after its optimistic flip
      setTimeout(function () { last = ""; push(); }, 700);
    });
  });
  setTimeout(push, 1500);
})(false);
