let g_MailboxItem, g_OfficeHostName, g_TimeOutMS = 4 * 60 * 1000 ; // 4 minutes;
let g_proto = "http", g_ServiceUrl = "//127.0.0.1";  // No I18N

Office.initialize = function (initialize) {
	g_MailboxItem = Office.context.mailbox.item;
	g_OfficeHostName = Office.context.mailbox.diagnostics.hostName;
};

async function getAttach() {
	return new Promise((resolve, reject) => {
		g_MailboxItem.getAttachmentsAsync(async (result) => {
			if (result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 0) {
				const attachments = result.value;
				try {
					const settled = await Promise.allSettled(
						attachments.map((attachment) => {
							return new Promise((res, rej) => {
								g_MailboxItem.getAttachmentContentAsync(attachment.id, (contentResult) => {
									if (contentResult.status === Office.AsyncResultStatus.Succeeded) {
										attachment.format = contentResult.value.format;
										attachment.content = contentResult.value.content;
										res(attachment);
									} else {
										rej(contentResult.error);
									}
								});
							});
						})
					);

					// Keep fulfilled results
					const successful = settled
						.filter((r) => r.status === "fulfilled")
						.map((r) => r.value);

					// Log rejected ones
					settled
						.filter((r) => r.status === "rejected")
						.forEach((r) => console.error("Attachment fetch failed:", r.reason));

					resolve(successful);
				}catch (err) {
					// This block only hits if Promise.allSettled itself blows up
					console.error("Unexpected error while fetching attachments:", err);
					reject(err);
				}
			} else {
				reject(result.error ?? new Error("No attachments found"));
			}
		});
	});
}

async function getAsyncWrapper(obj, param = null)
	return new Promise((resolve) => {
		const callback = (result) => {
			if (result.status === Office.AsyncResultStatus.Succeeded) {
				resolve(result.value);
			}
			else {
				console.error("Failed to get value:", result.error); // No I18N
				reject(result.error);
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
			from: await getAsyncWrapper(g_MailboxItem.from).catch(() => ""),
			to: await getAsyncWrapper(g_MailboxItem.to).catch(() => ""),
			cc: await getAsyncWrapper(g_MailboxItem.cc).catch(() => ""),
			bcc: await getAsyncWrapper(g_MailboxItem.bcc).catch(() => ""),
			subject: await getAsyncWrapper(g_MailboxItem.subject).catch(() => ""),
			body: await getAsyncWrapper(g_MailboxItem.body, Office.CoercionType.Text).catch(() => ""),
			attachments: await getAttach().catch(() => []),
			timestamp: Date.now()
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
