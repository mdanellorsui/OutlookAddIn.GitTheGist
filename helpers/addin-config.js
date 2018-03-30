function getConfig() {
	var config = {};
	
	config.gitHubUserName = Office.context.roamingSettings.get('gitHubUserName');
	config.defaultGistId = Office.context.roamingSettings.get('defaultGistId');

	// config.gitHubUserName = 'mstarrrsui';
	// config.defaultGistId = 12;

	return config;
}

function setConfig(config, callback) {
	Office.context.roamingSettings.set('gitHubUserName', config.gitHubUserName);
	Office.context.roamingSettings.set('defaultGistId', config.defaultGistId);
	
	Office.context.roamingSettings.saveAsync(callback);
}