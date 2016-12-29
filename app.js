'use strict';

const envx = require("envx");

const restify = require('restify');
const builder = require('botbuilder');
const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const expressSession = require('express-session');
const crypto = require('crypto');
const querystring = require('querystring');
const https = require('https');
const request = require('request');

//bot application identity
const MICROSOFT_APP_ID = envx("MICROSOFT_APP_ID");
const MICROSOFT_APP_PASSWORD = envx("MICROSOFT_APP_PASSWORD");

//oauth details
const AZUREAD_APP_ID = envx("AZUREAD_APP_ID");
const AZUREAD_APP_PASSWORD = envx("AZUREAD_APP_PASSWORD");
const AZUREAD_APP_REALM = envx("AZUREAD_APP_REALM");
const AUTHBOT_CALLBACKHOST = envx("AUTHBOT_CALLBACKHOST");
const AUTHBOT_STRATEGY = envx("AUTHBOT_STRATEGY");

//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3979, function () {
  console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat bot
console.log('started...')
console.log(MICROSOFT_APP_ID);
var connector = new builder.ChatConnector({
  appId: MICROSOFT_APP_ID,
  appPassword: MICROSOFT_APP_PASSWORD
});
var bot = new builder.UniversalBot(connector);
server.post('/api/messages', connector.listen());
server.get('/', restify.serveStatic({
  'directory': __dirname,
  'default': 'index.html'
}));
//=========================================================
// Auth Setup
//=========================================================

server.use(restify.queryParser());
server.use(restify.bodyParser());
server.use(expressSession({ secret: 'keyboard cat', resave: true, saveUninitialized: false }));
server.use(passport.initialize());

server.get('/login', function (req, res, next) {
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', customState: req.query.address, resourceURL: process.env.MICROSOFT_RESOURCE }, function (err, user, info) {
    console.log('login');
    if (err) {
      console.log(err);
      return next(err);
    }
    if (!user) {
      return res.redirect('/login');
    }
    req.logIn(user, function (err) {
      if (err) {
        return next(err);
      } else {
        return res.send('Welcome ' + req.user.displayName);
      }
    });
  })(req, res, next);
});

server.get('/api/OAuthCallback/',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login' }),
  (req, res) => {
    console.log('OAuthCallback');
    console.log(req);
    const address = JSON.parse(req.query.state);
    const magicCode = crypto.randomBytes(4).toString('hex');
    const messageData = { magicCode: magicCode, accessToken: req.user.accessToken, refreshToken: req.user.refreshToken, userId: address.user.id, name: req.user.displayName, email: req.user.preferred_username };
    
    var continueMsg = new builder.Message().address(address).text(JSON.stringify(messageData));
    console.log(continueMsg.toMessage());

    bot.receive(continueMsg.toMessage());
    res.send('Welcome ' + req.user.displayName + '! Please copy this number and paste it back to your chat so your authentication can complete: ' + magicCode);
});

passport.serializeUser(function(user, done) {
  done(null, user);
});

passport.deserializeUser(function(id, done) {
  done(null, id);
});

// Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
// For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
var realm = AZUREAD_APP_REALM; 
let oidStrategyv2 = {
  redirectUrl: AUTHBOT_CALLBACKHOST + '/api/OAuthCallback',
  realm: realm,
  clientID: AZUREAD_APP_ID,
  clientSecret: AZUREAD_APP_PASSWORD,
  identityMetadata: 'https://login.microsoftonline.com/' + realm + '/v2.0/.well-known/openid-configuration',
  skipUserProfile: false,
  validateIssuer: false,
  //allowHttpForRedirectUrl: true,
  responseType: 'code',
  responseMode: 'query',
  scope:['email', 'profile', 'offline_access', 'https://outlook.office.com/mail.read'],
  passReqToCallback: true
};

// Use the v1 endpoint (applications configured by manage.windowsazure.com)
// This works against Azure AD
let oidStrategyv1 = {
  redirectUrl: process.env.AUTHBOT_CALLBACKHOST +'/api/OAuthCallback',
  realm: process.env.MICROSOFT_REALM,
  clientID: process.env.MICROSOFT_CLIENT_ID,
  clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
  validateIssuer: false,
  //allowHttpForRedirectUrl: true,
  oidcIssuer: undefined,
  identityMetadata: 'https://login.microsoftonline.com/' + realm + '/.well-known/openid-configuration',
  skipUserProfile: true,
  responseType: 'code',
  responseMode: 'query',
  passReqToCallback: true
};

let strategy = null;
if ( AUTHBOT_STRATEGY == 'oidStrategyv1') {
  strategy = oidStrategyv1;
}
if ( AUTHBOT_STRATEGY == 'oidStrategyv2') {
  strategy = oidStrategyv2;
}

passport.use(new OIDCStrategy(strategy,
  (req, iss, sub, profile, accessToken, refreshToken, done) => {
    if (!profile.displayName) {
      return done(new Error("No oid found"), null);
    }
    // asynchronous verification, for effect...
    process.nextTick(() => {
      console.log('passport callback. access token:');
      console.log(accessToken);
      profile.accessToken = accessToken;
      console.log('passport callback. refreshtoken:');
      console.log(refreshToken);
      profile.refreshToken = refreshToken;

      return done(null, profile);
    });
  }
));

//=========================================================
// Bots Dialogs
//=========================================================
function login(session) {
  // Generate signin link
  const address = session.message.address;

  // TODO: Encrypt the address string
  const link = process.env.AUTHBOT_CALLBACKHOST + '/login?address=' + querystring.escape(JSON.stringify(address));
  

  var msg = new builder.Message(session) 
    .attachments([ 
        new builder.SigninCard(session) 
            .text("Please click this link to sign in first.") 
            .button("signin", link) 
    ]); 
  session.send(msg);
  builder.Prompts.text(session, "You must first sign into your account.");
}
bot.dialog('signin', [
  (session, results) => {
    console.log('signin callback: ' + results);
    session.endDialog();
  }
]);

bot.dialog('/', [
  (session, args, next) => {
    if (!(session.userData.userName && session.userData.accessToken && session.userData.refreshToken)) {
      session.send("Welcome! This bot retrieves the latest email for you after you login.");
      session.beginDialog('signinPrompt');
    } else {
      next();
    }
  },
  (session, results, next) => {
    if (session.userData.userName && session.userData.accessToken && session.userData.refreshToken) {
      // They're logged in
      console.log('refreshToken');
      console.log(session.userData.refreshToken);
      session.send("Welcome " + session.userData.userName + "! You are currently logged in. To quit, type 'quit'. To log out, type 'logout'. ");
      getUserLatestEmail(session.userData.accessToken,
        function (requestError, result) {
          if (result && result.value && result.value.length > 0) {
            const responseMessage = 'Hi ' + session.userData.userName +  ', Your latest email is: ' + result.value[0].Subject;

            console.log(responseMessage);
            session.send(responseMessage);
            getAccessTokenWithRefreshToken(session.userData.refreshToken, (err, body, res) => {

              if (err || body.error) {
                session.send("Something happened, " + err);
                session.endDialog();
              }else{
                console.log(body);
              }
              
            });
            
          }else{
            console.log('no user returned');
            if(requestError){
              console.error(requestError);
              getAccessTokenWithRefreshToken(session.userData.refreshToken, (err, body, res) => {

                if (err || body.error) {
                  session.send("Something happened, " + err);
                  session.endDialog();
                }else{
                  console.log(body);
                }
                
              });
            }
          }
        }
      );
      session.beginDialog('workPrompt');
    } else {
      session.endConversation("Goodbye.");
    }
  },
  (session, results) => {
    if (!session.userData.userName) {
      session.endConversation("Goodbye. You have been logged out.");
    } else {
      session.endConversation("Goodbye.");
    }
  }
]);

bot.dialog('workPrompt', [
  (session) => {
    builder.Prompts.text(session, "Type something to continue...");
  },
  (session, results) => {
    var prompt = results.response;
    if (prompt === 'logout') {
      session.userData.userName = null;
      //session.userData.userEmail = null;
      session.endDialog();
    } else if (prompt === 'quit') {
      session.endDialog();
    } else {
      session.replaceDialog('workPrompt');
    }
  }
]);

bot.dialog('signinPrompt', [
  (session, args) => {
    if (args && args.invalid) {
      // Re-prompt the user to click the link
      builder.Prompts.text(session, "please click the signin link.");
    } else {
      login(session);
    }
  },
  (session, results) => {
    //resuming
    console.log('resume: ' + results);
    session.userData.loginData = JSON.parse(results.response);
    if (session.userData.loginData && session.userData.loginData.magicCode && session.userData.loginData.accessToken) {
      session.beginDialog('validateCode');
    } else {
      session.replaceDialog('signinPrompt', { invalid: true });
    }
  },
  (session, results) => {
    if (results.response) {
      //code validated
      session.userData.userName = session.userData.loginData.name;
      session.endDialogWithResult({ response: true });
    } else {
      session.endDialogWithResult({ response: false });
    }
  }
]);

bot.dialog('validateCode', [
  (session) => {
    builder.Prompts.text(session, "Please enter the code you received or type 'quit' to end. ");
  },
  (session, results) => {
    const code = results.response;
    if (code === 'quit') {
      session.endDialogWithResult({ response: false });
    } else {
      if (code === session.userData.loginData.magicCode) {
        // Authenticated, save
        session.userData.accessToken = session.userData.loginData.accessToken;
        session.userData.refreshToken = session.userData.loginData.refreshToken;
        console.log('session.userData.accessToken:');
        console.log(session.userData.accessToken);
        console.log('session.userData.refreshToken:');
        console.log(session.userData.refreshToken);

        session.endDialogWithResult({ response: true });
      } else {
        session.send("hmm... Looks like that was an invalid code. Please try again.");
        session.replaceDialog('validateCode');
      }
    }
  }
]);
function getAccessTokenWithRefreshToken(refreshToken, callback){
  console.log('getAccessTokenWithRefreshToken');
  var data = 'grant_type=refresh_token' 
        + '&refresh_token=' + refreshToken
        + '&client_id=' + process.env.MICROSOFT_CLIENT_ID
        + '&client_secret=' + encodeURIComponent(process.env.MICROSOFT_CLIENT_SECRET) 

  var options = {
      method: 'POST',
      url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      body: data,
      // json: true,
      headers: { 'Content-Type' : 'application/x-www-form-urlencoded' }
  };

  // var options = {
  //   host: 'login.microsoftonline.com', //https://outlook.office.com/api/v2.0/me/messages
  //   path: '/common/oauth2/v2.0/token',
  //   method: 'POST',
  //   headers: {
  //     'Content-Type': 'application/x-www-form-urlencoded',
  //     'Content-Length': Buffer.byteLength(data)
  //   }
  // };

  // var postrequest = https.request(options, function (res) {
  //   res.setEncoding('utf8');
  //   res.on('data', function (chunk) {
  //       console.log('Response: ' + chunk);
  //   });
  // });

  // postrequest.write(data);
  // postrequest.end();
  request(options, function (err, res, body) {
      if (err) return callback(err, body, res);
      if (parseInt(res.statusCode / 100, 10) !== 2) {

          console.log(body);
          if (body.error) {
              return callback(new Error(res.statusCode + ': ' + (body.error.message || body.error)), body, res);
          }
          if (!body.access_token) {
              return callback(new Error(res.statusCode + ': refreshToken error'), body, res);
          }
          return callback(null, body, res);
      }
      callback(null, {
          accessToken: body.access_token,
          refreshToken: body.refresh_token
      }, res);
  }); 
}

function getUserLatestEmail(accessToken, callback) {
  console.log('getUserLatestEmail');
  console.log(accessToken);
  var options = {
    host: 'outlook.office.com', //https://outlook.office.com/api/v2.0/me/messages
    path: '/api/v2.0/me/MailFolders/Inbox/messages?$top=1',
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json',
      Authorization: 'Bearer ' + accessToken
    }
  };
  https.get(options, function (response) {
    var body = '';
    response.on('data', function (d) {
      body += d;
    });
    response.on('end', function () {
      var error;
      if (response.statusCode === 200) {
        callback(null, JSON.parse(body));
      } else {
        error = new Error();
        error.code = response.statusCode;
        error.message = response.statusMessage;
        // The error body sometimes includes an empty space
        // before the first character, remove it or it causes an error.
        body = body.trim();
        error.innerError = JSON.parse(body).error;
        callback(error, null);
      }
    });
  }).on('error', function (e) {
    callback(e, null);
  });
}