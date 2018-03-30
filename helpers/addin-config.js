function getConfig() {
	var config = {};
	
	// to force the addin to request the config settings
	//deleteConfig();

	config.gitHubUserName = Office.context.roamingSettings.get('gitHubUserName');
	config.defaultGistId = Office.context.roamingSettings.get('defaultGistId');

	// config.gitHubUserName = 'mdanellorsui';
	// config.defaultGistId = 'b7045a1aff44fd39302e532baa7365ef';

	return config;
}

function setConfig(config, callback) {
	Office.context.roamingSettings.set('gitHubUserName', config.gitHubUserName);
	Office.context.roamingSettings.set('defaultGistId', config.defaultGistId);
	
	Office.context.roamingSettings.saveAsync(callback);
}

function deleteConfig() {
	Office.context.roamingSettings.remove('gitHubUserName');
	Office.context.roamingSettings.remove('defaultGistId');
}