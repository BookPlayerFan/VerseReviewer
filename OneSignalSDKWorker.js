// Tapping a push should open the right tab. When the app is already running, iOS just brings it forward and does not
// reload it, so the page URL can't carry the choice. Tell the open app directly. This listener is added before OneSignal's
// own worker code so it always gets to run.
function findKey(o, key, depth) {
  if (!o || typeof o !== "object" || depth > 6) return null;
  if (typeof o[key] === "string") return o[key];
  for (var k in o) {
    var v = o[k];
    if (typeof v === "string" && v.charAt(0) === "{") { try { v = JSON.parse(v); } catch (e) { continue; } }
    var r = findKey(v, key, depth + 1);
    if (r !== null) return r;
  }
  return null;
}
self.addEventListener("notificationclick", function(event) {
  try {
    var n = event.notification, tab = findKey(n && n.data, "vrgo", 0);
    if (!tab && n && n.tag) {
      if (/^vr-new-/.test(n.tag)) tab = "new";
      else if (n.tag === "vr-daily" || /^vr-eve-/.test(n.tag)) tab = "review";
    }
    if (tab !== "new" && tab !== "review") return;
    event.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(function(list) {
      list.forEach(function(c) { try { c.postMessage({ type: "VR_GOTO", tab: tab }); } catch (e) {} });
    }));
  } catch (e) {}
});

// The page asks this worker to close delivered notifications (the worker is what showed them). "fresh" = new-card reminders, "daily" = the
// morning one, "all" = everything (the status panel's test button). The kind is the tag when it is ours, else the "vrkind" in the push's data.
self.addEventListener("message", function(event) {
  var m = event.data;
  if (!m || m.type !== "VR_CLOSE") return;
  var src = event.source;
  event.waitUntil(self.registration.getNotifications().then(function(list) {
    var out = { type: "VR_CLOSED", id: m.id, found: 0, closed: 0, tags: [] };
    list.forEach(function(n) {
      var kind = String(n.tag || "");
      if (!/^vr-/.test(kind)) {
        kind = "";
        try { var k = JSON.stringify(n.data || {}).match(/"vrkind":"(vr-[a-z]+)"/); if (k) kind = k[1]; } catch (e) {}
      }
      out.found++;
      out.tags.push(kind || "?");
      if (m.all === true || (m.fresh === true && /^vr-new/.test(kind)) || (m.daily === true && kind === "vr-daily")) { try { n.close(); out.closed++; } catch (e) {} }
    });
    try { if (src) src.postMessage(out); } catch (e) {}
  }, function() {
    try { if (src) src.postMessage({ type: "VR_CLOSED", id: m.id, found: 0, closed: 0, tags: [], failed: true }); } catch (e) {}
  }));
});

importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");
