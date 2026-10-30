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
    if (!tab && n && n.tag) {   // no vrgo in the data (an older push): the topic says which tab
      if (n.tag === "vr-new") tab = "new";
      else if (n.tag === "vr-daily") tab = "review";
    }
    if (tab !== "new" && tab !== "review") return;
    event.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(function(list) {
      list.forEach(function(c) { try { c.postMessage({ type: "VR_GOTO", tab: tab }); } catch (e) {} });
    }));
  } catch (e) {}
});

importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");
