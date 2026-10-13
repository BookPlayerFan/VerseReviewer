importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");

// Sets the home-screen badge from the number each push carries (key "vrbadge"). OneSignal still shows the notification itself.
function findBadge(o, depth) {
  if (!o || typeof o !== "object" || depth > 6) return null;
  if (typeof o.vrbadge === "number") return o.vrbadge;
  for (var k in o) {
    var v = o[k];
    if (typeof v === "string" && v.charAt(0) === "{") { try { v = JSON.parse(v); } catch (e) { continue; } }
    var r = findBadge(v, depth + 1);
    if (r !== null) return r;
  }
  return null;
}
self.addEventListener("push", function(event) {
  try {
    var n = event.data ? findBadge(event.data.json(), 0) : null;
    if (n === null || !self.navigator) return;
    if (n > 0 && self.navigator.setAppBadge) event.waitUntil(self.navigator.setAppBadge(n));
    else if (self.navigator.clearAppBadge) event.waitUntil(self.navigator.clearAppBadge());
  } catch (e) {}
});
