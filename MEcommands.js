let g_MailboxItem, g_OfficeHostName, officeHostName;

  OFFICE_HOST_NAMES = {
    OUTLOOK_CLIENT: "Outlook",
    OUTLOOK_WEB_ACCESS: "OutlookWebApp"
  };

Office.initialize = function (initialize) {
  g_MailboxItem = Office.context.mailbox.item;
  g_OfficeHostName = Office.context.mailbox.diagnostics.hostName;
const hostName = Office.context.mailbox.diagnostics.hostName;
  console.log("Outlook hostName:", hostName);
  officeHostName = Office.context.mailbox.diagnostics.hostName;
};

async function getAttach(){
  return new Promise((resolve) => {
    g_MailboxItem.getAttachmentsAsync(async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 0) {
        const attachments = result.value;
        try {
          const enriched = await Promise.all(
            attachments.map((attachment) => {
              return new Promise((res, rej) => {
                g_MailboxItem.getAttachmentContentAsync(attachment.id, (contentResult) => {
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

async function getAsyncWrapper(obj, param = null) {
  return new Promise((resolve) => {
    const callback = (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        console.error("Failed to get value:", result.error);
        resolve("");
      }
    };

    if (param !== null) {
      obj.getAsync(param, callback);
    } else {
      obj.getAsync(callback);
    }
  });
}

async function checkAvailableAgentPort() {
  let resolvedPort = null;
  const candidatePorts = [7212, 7412, 7612, 7812];
  for (let port of candidatePorts) {
    const url = `http://127.0.0.2:${port}/OutLook/MEDLP/v1.0/PortCheck`;
    try {
      const response = await fetch(url, { method: "GET", mode: "cors" });
      if (response.ok) {
        console.log("Port alive:", port);
        resolvedPort = port;
        break;
      } else {
        console.error(`Port ${port} responded with status ${response.status}`);
      }
    } catch (err) {
      console.error(`Port ${port} not available:`, err.message);
    }
  }
  return resolvedPort; 
}

async function eventValidator(event) {
  try {
    let agentPort = await checkAvailableAgentPort();
	if (!agentPort) {
    	console.error("No port available.");
    	if (event) {
      		event.completed({ allowEvent: true });
    	}
  	}
    const emailData = {
      from: await getAsyncWrapper(g_MailboxItem.from),
      to: await getAsyncWrapper(g_MailboxItem.to),
      cc: await getAsyncWrapper(g_MailboxItem.cc),
      bcc: await getAsyncWrapper(g_MailboxItem.bcc),
      subject: await getAsyncWrapper(g_MailboxItem.subject),
      body: await getAsyncWrapper(g_MailboxItem.body, Office.CoercionType.Text),
      attachments: await getAttach()
    };

    const url = `http://127.0.0.2:${agentPort}/OutLook/MEDLP/v1.0/Process`;
    const response = await fetch(url, {
      method: "POST",
      headers: {
      "Content-Type": "application/json;charset=utf-8",
      "Access-Control-Request-Method": "POST"
      },
      body: JSON.stringify(emailData, null, 2)
    });

    const result = await response.json();
    console.log("Response from EDLP :", result);
	
	if(result.allowEvent) {
		event.completed({ allowEvent: true });
	}
	else {
		event.completed({ allowEvent: false });
	}
  } catch (error) {
    console.error("Error in generate:", error);
    event.completed({ allowEvent: true });
  }
}

function onMessageSendHandler(event) {
  console.log("OnSend triggered.");
  try {
	  // Add-in runs only on Windows with Outlook Mailbox API v1.8+
	  if("Win32" === navigator.platform && Office.context.requirements.isSetSupported("Mailbox", 1.8) && officeHostName !== "Outlook") {
    	eventValidator(event);
	  }
	  else {
		console.error("Add in not supported");
		event.completed({ allowEvent: true });
	  }
	  
      g_MailboxItem.notificationMessages.addAsync("unsupported", {
        type: "errorMessage",
        message: "Not supported"
        });
	  
  } catch (err) {
    console.error("Error in OnSend:", err);
    event.completed({ allowEvent: true });
  }
}
