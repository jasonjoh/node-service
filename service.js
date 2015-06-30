var config = 'service-config.json';

var fs = require('fs');
var nconf = require('nconf');
var adal = require('adal-node');

var AuthenticationContext = adal.AuthenticationContext;

// Logging function for ADAL
function turnOnLogging() {
  var log = adal.Logging;
  log.setLoggingOptions(
  {
    level : log.LOGGING_LEVEL.VERBOSE,
    log : function(level, message, error) {
      console.log(message);
      if (error) {
        console.log(error);
      }
    }
  });
}

// Load configuration file and make sure we have client ID and certificate
nconf.env();
nconf.file({file: config});

var client_id = nconf.get('client_id');
if (client_id === undefined || client_id === "") {
  console.log('Client ID is required. Please enter your app\'s client ID in the client_id value in service-config.json.');
  process.exit(1);
}

var cert_file = nconf.get('cert_file');
if (cert_file === undefined || cert_file === "") {
  console.log('App certificate is required. Please enter the path to your app\'s PEM certificate file in the cert_file value in service-config.json.');
  process.exit(1);
}

var thumbprint = nconf.get('cert_thumbprint');
if (thumbprint === undefined || thumbprint === "") {
  console.log('Certificate thumbprint is required. Please enter the thumbprint for your app\'s certificate in the cert_thumbprint value in service-config.json.');
  process.exit(1);
}

var tenant = nconf.get('tenant');
if (tenant === undefined || tenant === "") {
  console.log('Tenant is required. Please enter the domain name for your Office 365 tenant in the tenant value in service-config.json.');
  process.exit(1);
}

console.log('Loaded configuration:');
console.log('  Client ID: ' + client_id);
console.log('  Certificate file: ' + cert_file);
console.log('  Certificate thumbprint: ' + thumbprint);
console.log('  Tenant: ' + tenant);

// Get the token

// Turn on ADAL logging
turnOnLogging();

var authorityUrl = 'https://login.microsoftonline.com/' + tenant;
var resource = 'https://outlook.office365.com';
var private_key = fs.readFileSync(cert_file, {encoding: 'utf8'});

var context = new AuthenticationContext(authorityUrl);

context.acquireTokenWithClientCertificate(resource, client_id, private_key, thumbprint, function (error, tokenResponse) {
  if (error) {
    console.log('ERROR acquiring token: ' + error.stack);
  } else {
    console.log('TOKEN: ' + JSON.stringify(tokenResponse));
  }
});