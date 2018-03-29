var config;
var btnEvent;

// The initialize function must be run each time a new page is loaded
Office.initialize = function(reason) {
	config = getConfig();
}

// Add any ui-less function here
function showError(error) {
	Office.context.mailbox.item.notificationMessages.replaceAsync('progress',  {
		type: 'progressIndicator',
		message: error
	});
}

var settingsDialog;



function insertDefaultGist(event) {

	Office.context.mailbox.item.notificationMessages.replaceAsync("progress", {
		type: "progressIndicator",
		message : "MAAARK!!!!!!!! config:" + JSON.stringify(config)
	  });

	  Office.context.mailbox.item.notificationMessages.replaceAsync("github-error", {
		type: "informationalMessage",
		message: "WHERE AM I????:" + JSON.stringify(config),
		icon : "iconid",
        persistent: false
	  });


	 
	// check if the add-in has been configured
	if ( config && config.defaultGistId) {
		// Get the default Gist content and insertDefaultGist
		Office.context.mailbox.item.notificationMessages.replaceAsync("error", {
			type: "errorMessage",
			message : "GETTING GIST!!!!"
		  });
		try {
			getGist(config.defaultGistId, function(gist, error) {
				if (gist) {
					buildBodyContent(gist, function(content, error) {
						if (content) {
							Office.context.mailbox.item.body.setSelectedDataAsync(content,
							{coercionType: Office.CoercionType.Html}, function(result) {
								event.completed();
							});
						} else {
							showError(error);
							event.completed();
						}
					});
				} else {
					showError(error);
					event.completed();
				}
			});
		} catch (err) {
			showError(err);
			event.completed();
		}
	} else {
		// save the event object so we can finish up later
		try {
			Office.context.mailbox.item.notificationMessages.replaceAsync("progress", {
				type: "progressIndicator",
				message : "MIKE!!!!!!!! Now we are down here."
			  });
			btnEvent = event;
			// Not Configured yet, display settings diaLOG WITH
			// WARN=1 TO display warning
			var url = new URI('../settings/dialog.html?warn=1').absoluteTo(window.location).toString();
			var dialogOptions = {width:20, height: 40, displayInIframe: false};
			
			Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
				settingsDialog = result.value;
				settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
				settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
			});
	
		}
		catch (err) {
			showError(err);
			event.completed();
		}
	}
}

function receiveMessage(message) {
	config = JSON.parse(message.message);
	setConfig(config, function(result) {
		settingsDialog.close();
		settingsDialog = null;
		btnEvent.completed();
		btnEvent = null;
	});
}

function dialogClosed(message) {
	settingsDialog = null;
	btnEvent.completed();
	btnEvent = null;
}