function getOAuthService() {
  return OAuth2.createService('NetSuite')
    .setAuthorizationBaseUrl('https://XXXXXXX.app.netsuite.com/app/login/oauth2/authorize.nl')
    .setTokenUrl('https://XXXXXX.suitetalk.api.netsuite.com/services/rest/auth/oauth2/v1/token')
    .setClientId('ID')
    .setClientSecret('SECRET')
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('rest_webservices restlets')
    .setParam('redirect_uri', 'https://script.google.com/macros/d/XXXXXXXXXXXXXXX/usercallback')
    .setParam('response_type', 'code');
}
function authCallback(request) {
  var oauthService = getOAuthService();
  var isAuthorized = oauthService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}
function authorize() {
  var oauthService = getOAuthService();
  if (!oauthService.hasAccess()) {
    var authorizationUrl = oauthService.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: ' + authorizationUrl);
  } else {
    Logger.log('Already authorized');
  }
}

function resetAuthorization() {
  var oauthService = getOAuthService();
  oauthService.reset();  // This will clear any stored tokens
  Logger.log('Authorization reset. Run the authorization process again.');
}