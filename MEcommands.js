let g_MailboxItem, g_OfficeHostName, g_TimeOutMS = 4 * 60 * 1000 ; // 4 minutes;
let g_proto = "http", g_ServiceUrl = "//127.0.0.1";  // No I18N

Office.initialize = function (initialize) {
	g_MailboxItem = Office.context.mailbox.item;
	g_OfficeHostName = Office.context.mailbox.diagnostics.hostName;
};

async function getAttach() {
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
									}
									else {
										console.error("Failed to get attachment content:", contentResult.error); // No I18N
										res(attachment);
									}
								});
							});
						})
					);
					resolve(enriched);
				}
				catch (err) {
					console.error("Error while fetching attachment contents:", err); // No I18N
					resolve(attachments);
				}
			}
			else {
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
			}
			else {
				console.error("Failed to get value:", result.error); // No I18N
				resolve("");
			}
		};

		if (param !== null) {
			obj.getAsync(param, callback);
		}
		else {
			obj.getAsync(callback);
		}
	});
}

async function checkAvailableAgentPort() {
	const candidatePorts = [7212, 7412, 7612];
	const checks = candidatePorts.map(port =>
		fetch(`${g_proto}:${g_ServiceUrl}:${port}/OutLook/MEDLP/v1.0/PortCheck`, {
			method: "GET", // No I18N
			mode: "cors" // No I18N
		})
			.then(response => {
				if (response.ok) {
					console.log("Port alive:", port);
					return port;
				}
				throw new Error(`Port ${port} responded with status ${response.status}`);
			})
			.catch(err => {
				console.error(`Port ${port} not available:`, err.message);
				throw err;
			})
	);

	try {
		return await Promise.any(checks);
	}
	catch {
		console.error("No available port found."); // No I18N
		return null;
	}
}

async function eventValidator(event) {
	try {
		let agentPort = await checkAvailableAgentPort();
		if (!agentPort) {
			console.error("No port available."); // No I18N
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

		const url = `${g_proto}:${g_ServiceUrl}:${agentPort}/OutLook/MEDLP/v1.0/Process`;

		const timeOutCallback = new Promise(resolve =>
			setTimeout(() => resolve({ allowEvent: true }), g_TimeOutMS)
		);

		const request = fetch(url, {
			method: "POST",  // No I18N
			headers: { "Content-Type": "application/json;charset=utf-8" },  // No I18N
			body: JSON.stringify(emailData)
		})
			.then(r => r.json())
			.then(res => ({ allowEvent: !!res.allowEvent }))
			.catch(() => ({ allowEvent: true }));

		const result = await Promise.race([timeOutCallback, request]);
		console.log("Response from EDLP :", result);
		event.completed(result);
	}
	catch (error) {
		console.error("Error in generate:", error); // No I18N
		event.completed({ allowEvent: true });
	}
}

function main(event) {
	console.log("OnSend triggered.");
	try {
		// Add-in runs only on Windows with new Outlook and Mailbox API v1.8+
		if ("Win32" === navigator.platform && Office.context.requirements.isSetSupported("Mailbox", 1.8) && g_OfficeHostName === "newOutlookWindows") {
			eventValidator(event);
		}
		else {
			console.error("Add in not supported"); // No I18N
			/* g_MailboxItem.notificationMessages.addAsync("Unsupported", {
				type: "errorMessage",
				message: "Addin does support"
			}); */
			event.completed({ allowEvent: true });
		}
	}
	catch (err) {
		console.error("Error in OnSend:", err); // No I18N
		event.completed({ allowEvent: true });
	}
}
