Office.initialize = function (a) {
  mailboxItem = Office.context.mailbox.item;
  officeHostName = Office.context.mailbox.diagnostics.hostName;
};
function onMessageAttachmentsChanged(eventArgs) {
    console.log("Attachment change event triggered.");

    const item = Office.context.mailbox.item;

    item.getAttachmentsAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to get attachments:", result.error);
            return;
        }

        const attachments = result.value;

        if (attachments.length === 0) {
            console.log("No attachments found.");
            return;
        }

        attachments.forEach((attachment, index) => {
            console.log(`\nAttachment #${index + 1}:`);
            console.log(`  Name: ${attachment.name}`);
            console.log(`  ID: ${attachment.id}`);
            console.log(`  Content Type: ${attachment.contentType}`);
            console.log(`  Attachment Type: ${attachment.attachmentType}`);

            // Now get the content of the attachment
            item.getAttachmentContentAsync(attachment.id, function (contentResult) {
                if (contentResult.status === Office.AsyncResultStatus.Succeeded) {
                    const content = contentResult.value;

                    console.log(`  Format: ${content.format}`);

                    switch (content.format) {
                        case Office.MailboxEnums.AttachmentContentFormat.Base64:
                            console.log("  Base64 Content:", content.content);
                            break;
                        case Office.MailboxEnums.AttachmentContentFormat.Eml:
                        case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
                            console.log("  Text Content:", content.content);
                            break;
                        default:
                            console.warn("  Unknown format");
                    }
                } else {
                    console.error("Failed to get attachment content:", contentResult.error);
                }
            });
        });
    });
}
function main(a) {
  console.log("hello main");
    a.completed({ allowEvent: !0 });
}
function mainHandleAttachments(a) {
  console.log("hello mainHandleAttachments");
    onMessageAttachmentsChanged(a);
    a.completed({ allowEvent: !0 });
}
function onMessageSendHandler(a) {
  console.log("hello main");
    a.completed({ allowEvent: !0 });
}
function hellow(){
  console.log("hellow world");
}
