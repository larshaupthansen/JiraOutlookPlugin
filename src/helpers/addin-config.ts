function getConfig() {
  var config: any = {};

  config.jiraUrl = Office.context.roamingSettings.get('jiraUrl');  

  return config;
}

function setConfig(config, callback) {
  Office.context.roamingSettings.set('jiraUrl', config.jiraUrl);
  Office.context.roamingSettings.saveAsync(callback);
}