/// <reference path="typings/node/node.d.ts"/>
var config = 'service-config.json';

var fs = require('fs');
var nconf = require('nconf');
var adal = require('adal-node');
var outlook = require('node-outlook');

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

// Function that gets a user's email
function GetUserEmail(user, accessToken) {
  
  // Uncomment this to enable tracing to the console
  //outlook.base.setTraceFunc(console.log);
  outlook.base.setFiddlerEnabled(true);
  
  var getMessages = outlook.base.apiEndpoint() + '/Users/' + user + '/Messages';
  
  var queryParams = {
    '$select': 'Subject,DateTimeReceived,From',
    '$orderby': 'DateTimeReceived',
    '$top': 5
  };
  
  // Option 1: Use makeApiCall to implement a GET
  // GET /Users/allieb@contoso.com/Messages?$select=Subject,DateTimeReceived,From&$orderby=DateTimeReceived&$top=5
  outlook.base.makeApiCall({url: getMessages, token: accessToken, query: queryParams}, function(error, response) {
    console.log('');
    if (error) {
      console.log('makeApiCall(GET) returned an error: ' + error);
    }
    if (response) {
      console.log('makeApiCall(GET) returned a response: ' + response.statusCode);
      console.log('Response body: ' + JSON.stringify(response.body, null, 2));
    }
    console.log('');
  });
  
  // Option 2: Use the getMessages function to do a GET
  outlook.mail.getMessages({token: accessToken, user: user, odataParams: queryParams}, function(error, response) {
    console.log('');
    if (error) {
      console.log('getMessages returned an error: ' + error);
    }
    if (response) {
      console.log('getMessages returned ' + response.value.length + ' messages.');
      response.value.forEach(function(message) {
        console.log('  Subject: ' + message.Subject);
      });
    }
    console.log('');
  });
}

function CreateUserEmail(user, accessToken) {
  var newMsg = {
    'Subject': "Created by Node Service app",
    'Importance': 'Low',
    'Body': {
      'ContentType': 'Text',
      'Content': 'This is an automated email created by the Node service.'
    },
    'ToRecipients': [
      {
        'EmailAddress': {
          'Address': 'test@example.com'
        }
      }
    ]
  };
  
  var putMessage = outlook.base.apiEndpoint() + '/Users/' + user + '/folders/drafts/messages';
  
  // POST /Users/allieb@contoso.com/folders/drafts/messages
  outlook.base.makeApiCall({url: putMessage, token: accessToken, method: 'POST', payload: newMsg}, function(error, response) {
    console.log('');
    if (error) {
      console.log('makeApiCall(POST) returned an error: ' + error);
    }
    if (response) {
      console.log('makeApiCall(POST) returned a response:' + response.statusCode);
      console.log('Response body: ' + JSON.stringify(response.body, null, 2));
    }
    console.log('');
  });
}

// Function that loops through users and gets their email
function GetUserEmails(accessToken) {
  var users = nconf.get('users');
  if (users === undefined || users.length <= 0) {
    console.log('No users specified. Please add users to the user value in service-config.json.');
  }
  
  users.forEach(function(user) {
    console.log('Getting mail for ' + user);
    GetUserEmail(user, accessToken);
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
console.log('');

// Get the token

// Uncomment this to turn on ADAL logging
//turnOnLogging();

var authorityUrl = 'https://login.microsoftonline.com/' + tenant;
var resource = 'https://outlook.office365.com';
var private_key = fs.readFileSync(cert_file, {encoding: 'utf8'});

var context = new AuthenticationContext(authorityUrl);

context.acquireTokenWithClientCertificate(resource, client_id, private_key, thumbprint, function (error, tokenResponse) {
  if (error) {
    console.log('ERROR acquiring token: ' + error.stack);
  } else {
    GetUserEmails(tokenResponse.accessToken);
  }
});