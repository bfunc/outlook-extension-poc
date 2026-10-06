/* OnMessageSend handler (Smart Alerts). Plain JS, no imports: classic Outlook on Windows runs this
   file in a JavaScript-only runtime. Acts only on messages where the task pane was opened
   (sessionData "pocUsed"); never blocks sending - on any error the message still goes out. */

var POC_API = "__POC_API__"; // replaced from VITE_API_BASE in .env at build time
var POC_SUBJECT_SUFFIX = " Outlook extension PoC";
var POC_DEBUG = false; // send step beacons to the server log (GET .../health?step=...)

function pocTrace(step) {
  if (!POC_DEBUG) return;
  try { fetch(POC_API + "/health?step=" + encodeURIComponent(step) + "&t=" + Date.now(), { cache: "no-store" }); } catch (e) {}
}

function pocCall(fn) {
  return new Promise(function (resolve, reject) {
    fn(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
      else reject(r.error);
    });
  });
}

// sessionData first (this compose session), then the custom properties saved with the draft.
function pocSession(key) {
  var item = Office.context.mailbox.item;
  return pocCall(function (cb) { item.sessionData.getAsync(key, cb); })
    .catch(function () { return ""; })
    .then(function (v) {
      if (v) return v;
      return pocCall(function (cb) { item.loadCustomPropertiesAsync(cb); })
        .then(function (props) { return props.get(key) || ""; })
        .catch(function () { return ""; });
    });
}

function pocEmails(list) {
  return (list || []).map(function (r) { return r.emailAddress; }).filter(Boolean);
}

function pocBuildPayload(itemIds) {
  var item = Office.context.mailbox.item;
  return Promise.all([
    pocCall(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); }),
    pocCall(function (cb) { item.to.getAsync(cb); }),
    pocCall(function (cb) { item.cc.getAsync(cb); }),
    pocCall(function (cb) { item.bcc.getAsync(cb); })
  ]).then(function (v) {
    return {
      itemIds: itemIds,
      email: v[0],
      recipients: pocEmails(v[1]).concat(pocEmails(v[2]), pocEmails(v[3]))
    };
  });
}

function pocAppendSubject() {
  var item = Office.context.mailbox.item;
  return pocCall(function (cb) { item.subject.getAsync(cb); }).then(function (subject) {
    subject = subject || "";
    if (subject.slice(-POC_SUBJECT_SUFFIX.length) === POC_SUBJECT_SUFFIX) return;
    return pocCall(function (cb) { item.subject.setAsync(subject + POC_SUBJECT_SUFFIX, cb); });
  });
}

function onMessageSendHandler(event) {
  pocTrace("handler:start");
  var done = false;
  function finish() {
    if (done) return;
    done = true;
    pocTrace("completed");
    event.completed({ allowEvent: true });
  }
  // Never hold the send for long: allow it after 20 s whatever happens.
  setTimeout(finish, 20000);
  pocSession("pocUsed")
    .then(function (used) {
      pocTrace("used:" + used);
      if (used !== "1") return;
      return Promise.all([pocSession("pocItemIds"), pocSession("pocSubmitted")]).then(function (s) {
        var itemIds = [];
        try { itemIds = JSON.parse(s[0] || "[]"); } catch (e) { itemIds = []; }
        pocTrace("session:" + s.join("|"));
        return pocAppendSubject().then(function () {
          pocTrace("subject:ok");
          if (s[1] === "1") return; // already posted from the task pane's Send button
          return pocBuildPayload(itemIds).then(function (payload) {
            return fetch(POC_API + "/submit", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(payload)
            });
          });
        });
      });
    })
    .then(finish, finish);
}

pocTrace("script:loaded");
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
pocTrace("associated");
// Outlook on the web: without an Office.onReady call the event runtime in commands.html never got the
// OnMessageSend event (the send hung on "taking longer than expected"). Classic Outlook's JS-only
// runtime does not run onReady callbacks, which is fine - the handler needs nothing from it.
Office.onReady(function () { pocTrace("ready"); });
