'use strict';

const restify = require('restify');
const builder = require('botbuilder');
const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const crypto = require('crypto');
const querystring = require('querystring');

//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3979, function () {
  console.log('%s listening to %s', server.name, server.url);
});

// Create chat bot
var connector = new builder.ChatConnector({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD
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
server.use(passport.initialize());

server.get('/login', function (req, res, next) {
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', state: req.query.address }, function (err, user, info) {
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

    // TODO: decrypt
    const address = JSON.parse(req.query.state);
    const magicCode = crypto.randomBytes(4).toString('hex');

    const messageData = { magicCode: magicCode, authCode: req.query.code, userId: address.user.id, name: req.user.displayName, email: req.user.email };

    bot.receive(continueMsg.toMessage());
    res.send('Welcome ' + req.user.displayName + '! Please copy this number and paste it back to your chat so your authentication can complete: ' + code);
  });

passport.serializeUser(function (user, done) {
  done(null, user);
});
passport.deserializeUser(function (id, done) {
  done(null, id);
});

// Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
// For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
var realm = process.env.MICROSOFT_REALM;
let oidStrategyv2 = {
  callbackURL: process.env.AUTHBOT_CALLBACKHOST + '/api/OAuthCallback',
  realm: realm,
  clientID: process.env.MICROSOFT_APP_ID,
  clientSecret: process.env.MICROSOFT_APP_PASSWORD,
  identityMetadata: 'https://login.microsoftonline.com/' + realm + '/v2.0/.well-known/openid-configuration',
  skipUserProfile: true,
  responseType: 'code',
  responseMode: 'query',
  scope: ['email', 'profile'],
  passReqToCallback: true
};

// Use the v1 endpoint (applications configured by manage.windowsazure.com)
// This works against Azure AD
let oidStrategyv1 = {
  callbackURL: process.env.AUTHBOT_CALLBACKHOST + '/api/OAuthCallback',
  realm: process.env.MICROSOFT_REALM,
  clientID: process.env.MICROSOFT_CLIENT_ID,
  clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
  oidcIssuer: undefined,
  identityMetadata: 'https://login.microsoftonline.com/common/.well-known/openid-configuration',
  skipUserProfile: true,
  responseType: 'code',
  responseMode: 'query',
  passReqToCallback: true
};

let strategy = null;
if (process.env.AUTHBOT_STRATEGY == 'oidStrategyv1') {
  strategy = oidStrategyv1;
  console.log('using v1');
}
if (process.env.AUTHBOT_STRATEGY == 'oidStrategyv2') {
  strategy = oidStrategyv2;
  console.log('using v2');
}

passport.use(new OIDCStrategy(strategy,
  (req, iss, sub, profile, accessToken, refreshToken, done) => {
    console.log('strategy returned');
    console.log('accessToken: ' + accessToken);
    console.log('refreshToken: ' + refreshToken);

    if (!profile.email) {
      console.log('no profile email');
      return done(new Error("No email found"), null);
    }
    // asynchronous verification, for effect...
    process.nextTick(() => {
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
  builder.Prompts.text(session, "Please signin: " + link);
}

bot.dialog('/', [
  (session, args, next) => {
    console.log(session.userData.userEmail);

    if (!session.userData.userEmail) {
      session.beginDialog('signinPrompt');
    } else {
      next();
    }
  },
  (session, results, next) => {
    if (session.userData.userEmail) {
      // They're logged in
      //var accessToken = session.privateConversationData.accessToken;
      session.send("Welcome " + session.userData.userEmail + "! You are currently logged in. To quit, type 'quit'. To log out, type 'logout'. ");
      session.beginDialog('workPrompt');
    } else {
      session.endConversation("Goodbye.");
    }
  },
  (session, results) => {
    if (!session.userData.userEmail) {
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
      session.userData.userEmail = null;
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
    console.log('signinPrompt');
    if (args && args.invalid) {
      // Re-prompt the user to click the link
      builder.Prompts.text(session, "please click the signin link.");
    } else {
      if (session.userData.refreshToken) {
        // TODO: Authorization
        //session.sendTyping();
        //get access token from refresh token
      } else {
        login(session);
      }
    }
  },
  (session, results) => {
    // Resume with "signin?authCode=<authCode>&code=<code>&name=<name>"
    console.log('resumed result: ' + results.response);

    session.userData.loginData = JSON.parse(results.response);
    if (session.userData.loginData && session.userData.loginData.magicCode && session.userData.loginData.authCode) {
      session.beginDialog('validateCode');
    } else {
      session.replaceDialog('signinPrompt', { invalid: true });
    }
  },
  (session, results) => {
    if (results.response) {
      //code validated
      session.userData.userName = session.userData.loginData.userName;
      session.userData.userEmail = session.dialogData.loginData.userEmail;
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
        session.userData.authCode = session.userData.loginData.authCode;
        // TODO: Authorize, then save
        session.endDialogWithResult({ response: true });
      } else {
        session.send("hmm... Looks like that was an invalid code. Please try again.");
        session.replaceDialog('validateCode');
      }
    }
  }
]);