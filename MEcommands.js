const SUPPORTED_PLATFORMS = ["MacIntel", "Win32"],
  AVAILABLE_SERVICE_PORTS = [4631, 4641, 4651],
  REQUEST_TIMEOUT = 3e5,
  CONTENT_HEADERS = { "Content-Type": "application/json;charset=utf-8" },
  POST_HEADERS = {
    "Content-Type": "application/json;charset=utf-8",
    "Access-Control-Request-Method": "POST"
  },
  OFFICE_HOST_NAMES = {
    OUTLOOK_CLIENT: "Outlook",
    OUTLOOK_WEB_ACCESS: "OutlookWebApp"
  };

let officeHostName,
  mailboxItem,
  agentWebSrvcPort,
  addinBrowserStorage = window.localStorage,
  serviceUrl = "http://127.0.0.1";

Office.initialize = function (a) {
  mailboxItem = Office.context.mailbox.item;
  officeHostName = Office.context.mailbox.diagnostics.hostName;
};

function request(a, b, c, d) {
  return new Promise((e, g) => {
    let f = new XMLHttpRequest();
    f.open(b, a);
    f.timeout = REQUEST_TIMEOUT;
    c && Object.keys(c).forEach(h => {
      f.setRequestHeader(h, c[h]);
    });
    f.onload = () => {
      200 == f.status ? e(f.response) : (error_str = "Non 200 response code received - " + f.statusText, g(error_str));
    };
    f.onerror = () => g(f.statusText);
    f.send(d);
  });
}

function getURL(a, b) {
  return `${serviceUrl}:${a}/OUTLOOK/v1/${b}`;
}

async function sendPortAvailablityRequestToAgentWebService(a) {
  a = getURL(a, "CheckBody");
  return await fetch(a, { mode: "cors", method: "HEAD" });
}

async function sendErrorToAgentWebService(a) {
  try {
    let b = getURL(agentWebSrvcPort, "error");
    return await fetch(b, { mode: "cors", method: "POST", headers: POST_HEADERS, body: a });
  } catch (b) {
    logError("Error occurred while sending error message to agent.");
    logError(b);
  }
}

async function sendMessageToAgentWebService(a) {
  logMessage("Sending message request to agent web service.");
  let b = getURL(agentWebSrvcPort, "CheckBody");
  return await request(b, "POST", CONTENT_HEADERS, a);
}

async function sendAttachmentCacheReqToAgentWebService(a) {
  logMessage("Sending attachment cache request to agent web service.");
  a = JSON.stringify(a);
  logMessage(a);
  let b = getURL(agentWebSrvcPort, "attachmentCache");
  return await request(b, "POST", CONTENT_HEADERS, a);
}

async function sendAttachmentCacheDeleteReqToAgentWebService(a) {
  logMessage("Sending attachment cache request to agent web service.");
  a = JSON.stringify({ attachmentIdList: a });
  logMessage(a);
  let b = getURL(agentWebSrvcPort, "attachmentCacheDelete");
  return await request(b, "POST", CONTENT_HEADERS, a);
}

async function deleteCachedAttachmentList(a) {
  try {
    res = await sendAttachmentCacheDeleteReqToAgentWebService(a);
    logMessage("Cache delete attachment response:");
    logMessage(res);
  } catch (b) {
    logError("Failed to delete the cache entry");
    logError(b);
  }
}

async function sendAttachmentCacheStatusReqToAgentWebService(a) {
  logMessage("Sending attachment cache status request to agent web service.");
  a = JSON.stringify({ attachmentIdList: a });
  logMessage(a);
  try {
    let b = getURL(agentWebSrvcPort, "getAttachmentCacheStatus"),
      c = await request(b, "POST", CONTENT_HEADERS, a),
      d = JSON.parse(c);
    logMessage("cache status response");
    logMessage(d);
    return d.cachedAttachmentIdList;
  } catch (b) {
    logMessage("Failed to get the attachment cache status. Exception: " + b.message);
    return [];
  }
}

function logMessage(a) { }

function logError(a) {
  console.error(a);
}

async function checkPortAvailability(a, b, c) {
  logMessage("Checking port availability for port: " + a);
  try {
    res = await sendPortAvailablityRequestToAgentWebService(a);
    res.ok ? (
      logMessage("Resolved port: " + a), b(a)
    ) : 400 === res.status ? c("Outlook channel is disabled.") : c("Non 200 response received. res = " + res);
  } catch (d) {
    c("Exception: " + d.message);
  }
}

async function resolveAgentWebSrvcPort(a) {
  logMessage("Resolving agent web service port.");
  let b = [];
  for (let c = 0; c < AVAILABLE_SERVICE_PORTS.length; c++) {
    b.push(new Promise((d, e) => {
      checkPortAvailability(AVAILABLE_SERVICE_PORTS[c], d, e);
    }));
  }
  await Promise.any(b)
    .then(c => {
      agentWebSrvcPort = c;
      logMessage("Resolved port: " + agentWebSrvcPort);
    })
    .catch(c => {
      logMessage("Couldn't resolve any port.");
      !0 === c instanceof AggregateError && logError(c.errors);
      a.completed({ allowEvent: !0 });
    });
}

async function isMailChangedAfterSendingToAgent(a, b) {
  logMessage("Checking if mail is changed after sending to agent.");
  var c = await outlookJSAPIHelper.getAttachmentIdList(mailboxItem);
  logMessage("Comparing for attachments");
  if (JSON.stringify(a) !== JSON.stringify(c)) {
    logMessage("Attachments changed. Mail is modified after sending.");
    return !0;
  }
  a = null;
  c = new MessageGenerator(mailboxItem);
  try {
    a = await c.generateMessageWithoutAttachment(mailboxItem.itemType);
    logMessage("Message object without attachment:");
    logMessage(a);
  } catch (d) {
    logError("Exception in generating message for post send verification. Allowing email to send.");
    logError(d);
    return !1;
  }
  logMessage("Comparing other message object details");
  return b.subject !== a.subject ||
    b.sender !== a.sender ||
    JSON.stringify(b.recipients) !== JSON.stringify(a.recipients) ||
    b.body !== a.body ||
    b.location !== a.location ||
    JSON.stringify(b.metaData) !== JSON.stringify(a.metaData) ? (
    logMessage("Mail is modified after sending."), !0
  ) : !1;
}

function isPluginSupported() {
  return !SUPPORTED_PLATFORMS.includes(navigator.platform) ||
    ("Win32" === navigator.platform && officeHostName === OFFICE_HOST_NAMES.OUTLOOK_CLIENT) ? !1 :
    Office.context.requirements.isSetSupported("Mailbox", 1.8) ? !0 : (
      mailboxItem.notificationMessages.addAsync("unsupported", {
        type: "errorMessage",
        message: "Not supported"
      }), !1
    );
}

function updateServiceUrlIfMacPlatform() {
  "MacIntel" == navigator.platform && (serviceUrl = "https://127.0.0.1");
}

async function getCachedAttachmentIdList() {
  let a = [];
  try {
    let b = await outlookJSAPIHelper.getAttachmentIdList(mailboxItem);
    0 < b.length && (
      logMessage("Checking Attachment List cached status:" + b),
      a = await sendAttachmentCacheStatusReqToAgentWebService(b)
    );
  } catch (b) {
    logMessage("Cached Attachment List status check function failed. Error:" + b.message);
  }
  return a;
}

async function isResponseAllow(a) {
  logMessage("response from agent web service");
  logMessage(a);
  return JSON.parse(a).allow ? !0 : !1;
}

async function validateBody(a) {
  if (!1 === isPluginSupported()) {
    logMessage("Plugin is not supported. Allowing email to sent.");
    a.completed({ allowEvent: !0 });
  } else {
    updateServiceUrlIfMacPlatform();
    await resolveAgentWebSrvcPort(a);
    var b = [],
      c = null,
      d = new MessageGenerator(mailboxItem);
    try {
      b = await getCachedAttachmentIdList();
      c = await d.generateMessage(mailboxItem.itemType, b);
    } catch (e) {
      logError(e);
      d = new AddInError(ERROR_TYPES.api_error, e.name, e.message + " error_code: " + e.code);
      errorMessage = new ErrorMessage();
      errorMessage.addError(d);
      sendErrorToAgentWebService(JSON.stringify(errorMessage));
    }
    d = JSON.stringify(c);
    logMessage("Json message : ");
    logMessage(d);
    try {
      let e = await outlookJSAPIHelper.getAttachmentIdList(mailboxItem);
      res = await sendMessageToAgentWebService(d);
      await isResponseAllow(res) ? (
        !1 === c.isBodyChangedInSuccessiveAPICalls && await isMailChangedAfterSendingToAgent(e, c) ? (
          logMessage("Mail is Modified after sending to agent."),
          a.completed({ allowEvent: !1 })
        ) : (
          0 < b.length && await deleteCachedAttachmentList(b),
          a.completed({ allowEvent: !0 })
        )
      ) : (
        logMessage("Response is not allowed."),
        a.completed({ allowEvent: !1 })
      );
    } catch (e) {
      logError(e);
      a.completed({ allowEvent: !0 });
    }
  }
}

async function mainHandleAttachments(a) {
  logMessage("Handling attachments");
  logMessage(a);
  if (!1 === isPluginSupported()) {
    logMessage("Plugin is not supported.");
  } else {
    updateServiceUrlIfMacPlatform();
    await resolveAgentWebSrvcPort(a);
    if (a.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
      a = a.attachmentDetails;
      a = new AttachmentCacheRequest(a.id, a.name, a.size);
      try {
        res = await sendAttachmentCacheReqToAgentWebService(a);
        logMessage("Successfully Cached the attachment");
      } catch (b) {
        logError(b);
      }
    } else if (a.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Removed) {
      let b = [];
      b.push(a.attachmentDetails.id);
      deleteCachedAttachmentList(b);
    }
  }
}

function main(a) {
  console.log("hello main");
    a.completed({ allowEvent: !0 });
}
function hellow(){
  console.log("hellow world");
}
