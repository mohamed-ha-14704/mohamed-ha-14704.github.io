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

function main(a) {
  console.log("hello main");
    a.completed({ allowEvent: !0 });
}
function mainHandleAttachments(a) {
  console.log("hello mainHandleAttachments");
    a.completed({ allowEvent: !0 });
}
function onMessageSendHandler(a) {
  console.log("hello main");
    a.completed({ allowEvent: !0 });
}
function hellow(){
  console.log("hellow world");
}
