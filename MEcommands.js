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
