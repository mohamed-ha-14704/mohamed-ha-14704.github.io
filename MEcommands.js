Office.initialize = function (a) {
  mailboxItem = Office.context.mailbox.item;
  officeHostName = Office.context.mailbox.diagnostics.hostName;
};

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
                    res(attachment);
                  }
                });
              });
            })
          );

          resolve(enriched);
        } catch (err) {
          console.error("Error while fetching attachment contents:", err);
          resolve(attachments);
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

async function validate(event) {
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
	
	if(result.action)
		event.completed({ allowEvent: true });
	else
		event.completed({ allowEvent: false });
  } catch (error) {
    console.error("Error in generate:", error);
    event.completed({ allowEvent: true });
  }
}

function onMessageSendHandler(event) {
  console.log("OnSend triggered.");
  try {
	  if("Win32" === navigator.platform && Office.context.requirements.isSetSupported("Mailbox", 1.8)) {
    	validate(event);
	  }
	  else {
		console.error("Add in not supported");
		event.completed({ allowEvent: true });
	  }
  } catch (err) {
    console.error("Error in OnSend:", err);
    event.completed({ allowEvent: true });
  }
}
