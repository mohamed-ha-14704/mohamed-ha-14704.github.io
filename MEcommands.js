Office.initialize = function (a) {
  mailboxItem = Office.context.mailbox.item;
  officeHostName = Office.context.mailbox.diagnostics.hostName;
};


function hellow()
{
}

async function onMessageAttachmentsChanged() {

    console.log("Attachment change event triggered.");
    const item = Office.context.mailbox.item;

    item.getAttachmentsAsync(result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to get attachments:", result.error);
            return;
        }
        console.log("Attachment go to read.");
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
    console.log("Attachment go to end.");
}

function mainHandleAttachments(a) {
  //console.log("hello attachment changes triggered");
  //onMessageAttachmentsChanged();
  a.completed({ allowEvent: !0 });
}

async function getAttach(){
  return new Promise((resolve) => {
    mailboxItem.getAttachmentsAsync(async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 0) {
        const attachments = result.value;

        try {
          const enriched = await Promise.all(
            attachments.map((attachment) => {
              return new Promise((res, rej) => {
                mailboxItem.getAttachmentContentAsync(attachment.id, (contentResult) => {
                  if (contentResult.status === Office.AsyncResultStatus.Succeeded) {
                    attachment.format = contentResult.value.format;
                    attachment.content = contentResult.value.content;
                    res(attachment);
                  } else {
                    console.error("Failed to get attachment content:", contentResult.error);
                    res(attachment); // still resolve with basic attachment info
                  }
                });
              });
            })
          );

          resolve(enriched);
        } catch (err) {
          console.error("Error while fetching attachment contents:", err);
          resolve(attachments); // fallback
        }

      } else {
        console.log("Failed to get attachments:", result.error);
        resolve([]);
      }
    });
  });
}

async function getParam(a, b) {
  return new Promise((resolve) => {
    a.getAsync(b, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        console.error("Failed to get subject:", result.error);
        resolve("");
      }
    });
  });
}

async function get(a) {
  return new Promise((resolve) => {
    a.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        console.error("Failed to get subject:", result.error);
        resolve("");
      }
    });
  });
}

async function generate(event) {
  try {
    const data = {
      from: await get(mailboxItem.from),
      to: await get(mailboxItem.to),
      cc: await get(mailboxItem.cc),
      bcc: await get(mailboxItem.bcc),
      subject: await get(mailboxItem.subject),
      body: await getParam(mailboxItem.body, Office.CoercionType.Text),
      attachments: await getAttach()
    };
    console.log("Email Metadata:", JSON.stringify(data, null, 2));

    const response = await fetch("http://127.0.0.1:7220/OutLook/MEDLP/v1.0/Process", {
      method: "POST",
      headers: {
      "Content-Type": "application/json;charset=utf-8",
      "Access-Control-Request-Method": "POST"
      },
      body: JSON.stringify(data, null, 2)
    });

    const result = await response.json();
    console.log("Response from native app:", result);
    // You can also use this data object for further processing
    // e.g., send to server, validate content, etc.
	if(result.action)
		event.completed({ allowEvent: true });
	else
		event.completed({ allowEvent: false });
  } catch (error) {
    console.error("Error in generate:", error);
    event.completed({ allowEvent: true }); // Allow send to proceed even on error
  }
}

function onMessageSendHandler(event) {
  console.log("OnSend triggered.");
  try {
    generate(event);
  } catch (err) {
    console.error("Error in OnSend:", err);
    event.completed({ allowEvent: true });
  }
}
