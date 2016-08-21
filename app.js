'user strict'

var restify = require('restify');
var builder = require('botbuilder');
var passport = require('passport');
var OIDCStrategy = require('passport-azure-ad').OIDCStrategy;

const querystring = require('querystring');
const AuthResultKey = "authResult";
const MagicNumberKey = "authMagicNumber";

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

// logged in users
var users = [];

var findByEmail = function (email, fn) {
  for (var i = 0, len = users.length; i < len; i++) {
    var user = users[i];
    if (user.email === email) {
      return fn(null, user);
    }
  }
  return fn(null, null);
};

server.use(restify.queryParser());
server.use(restify.bodyParser());
server.use(passport.initialize());

server.get('/login', function(req, res, next) {
  console.log('get login');
  console.log("passed in state: " + req.query.state);
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', state: req.query.state }, function(err, user, info) {
    console.log('login callback');
    if (err) { return next(err); }
    if (!user) { return res.redirect('/login'); }
    req.logIn(user, function(err) {
      if (err) { return next(err); }
      return res.send('Welcome ' + req.user.displayName);
    });
  })(req, res, next);
});

server.get('/api/OAuthCallback/',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', state: 'ritatest' }),
  function (req, res) {
    console.log('Returned from AzureAD.');
    console.log('token: ' + req.query.code);
    console.log('state: ' + req.query.state);

    // TODO: decrypt
    //var state = JSON.parse(decodeURIComponent(req.query.state));
    //console.log('decoded and parsed state: ' + state);
    // var addressId = state.addressId;
    // var conversationId = state.conversationId;
    var addressId = req.query.state.split('|')[0];
    var conversationId = req.query.state.split('|')[1];
    console.log('addressId: ' + addressId + '|' + 'conversationId: ' + conversationId);

    var token = req.query.code;
    var code = 'zhang';

  var address = { id: addressId,
     channelId: 'webchat',
     user: { id: '6MOefHwNrNo', name: '6MOefHwNrNo' },
     conversation: { id: conversationId },
     bot: { id: 'authbot', name: 'authbot' },
     serviceUrl: 'https://webchat.botframework.com',
     useAuth: true };

    var continueMsg = new builder.Message().address(address).text("signin?token=" + token + "&code=" + code + "&name=" + req.user.displayName);
    bot.receive(continueMsg.toMessage());
    res.send('Welcome ' + req.user.displayName);
  });

passport.serializeUser(function(user, done) {
    console.log('passport.serializeUser');
    done(null, user.email);
});
passport.deserializeUser(function(id, done) {
    console.log('passport.deserializeUser');
  findByEmail(id, function (err, user) {
    done(err, user);
  });
});

// Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
let oidStrategyv2 = {
    callbackURL: process.env.AUTHBOT_CALLBACKHOST + '/api/OAuthCallback',
    realm: 'common',
    clientID: process.env.MICROSOFT_APP_ID,
    clientSecret: process.env.MICROSOFT_APP_PASSWORD,
    identityMetadata: 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
    skipUserProfile: true,
    responseType: 'code',
    responseMode: 'query',
    scope: ['email', 'profile'],
    passReqToCallback: true
};

// Use the v1 endpoint (applications configured by manage.windowsazure.com)
let oidStrategyv1 = {
    callbackURL: process.env.AUTHBOT_CALLBACKHOST +'/api/OAuthCallback',
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
if ( process.env.AUTHBOT_STRATEGY == 'oidStrategyv1') {
  strategy = oidStrategyv1;
  console.log('using v1');
}
if ( process.env.AUTHBOT_STRATEGY == 'oidStrategyv2') {
  strategy = oidStrategyv2;
  console.log('using v2');
}

passport.use(new OIDCStrategy(strategy,
  function(req, iss, sub, profile, accessToken, refreshToken, done) {
    console.log('strategy returned');

    if (!profile.email) {
      console.log('no profile email');
      return done(new Error("No email found"), null);
    }
    // asynchronous verification, for effect...
    process.nextTick(function () {
      findByEmail(profile.email, function(err, user) {
        if (err) {
          return done(err);
        }
        if (!user) {
          // "Auto-registration"
          console.log('add new email');
          users.push(profile);
          return done(null, profile, {rita:'awesome'});
        }
        console.log('retrieve email');
        return done(null, user, {rita:'awesome'});
      });
    });
  }
));

//=========================================================
// Bots Dialogs
//=========================================================
function login(session) {
    // Generate signin link
    console.log(session.message);
    var addressId = session.message.address.id;
    var conversationId = session.message.address.conversation.id;
    // var resumptionCookie = JSON.stringify({addressId: addressId, conversationId:conversationId}); //JSON.stringify(session.message.address); 
    // console.log('creating resumptionCookie: ' + resumptionCookie);
    // resumptionCookie = encodeURIComponent(resumptionCookie);
    // TODO: encrypt
    var resumptionCookie = addressId + "|" + conversationId;

    var link = process.env.AUTHBOT_CALLBACKHOST + '/login?state=' + resumptionCookie;
    console.log(link);
    builder.Prompts.text(session, "Please signin: " + link);
}

bot.dialog('/', [
  function (session) {
      session.beginDialog('signinPrompt');
  },
  function (session, results) {
      if (results.response && results.username) {
          // They're logged in
          //var accessToken = session.privateConversationData.accessToken;
          session.send("Welcome " + results.username + "! You are logged in.");
      } else {
          session.endConversation("Goodbye");
      }
  }
]);

bot.dialog('signinPrompt', [
  function (session, args) {
    console.log('signinPrompt');
    if (args && args.invalid) {
      // Re-prompt the user to click the link
      builder.Prompts.text(session, "please click the signin link.");
    } else {
        if (session.userData.refreshToken) {
          //session.sendTyping();
          // TODO: get access token from refresh token
        } else {
            login(session);
        }
    }
  },
  function (session, results) {
    // Resume with "signin?token=<token>&code=<code>&name=<name>"
    console.log('resumed result: ' + results.response);

    var parts = results.response.split('?');
    if (parts.length == 2 && parts[0] == 'signin') {
      var params = parts[1].split('&');
      if (params.length == 3 && params[0].indexOf('token') > -1) {
          // Store token
          var token = params[0].split('=')[1];
          var code = params[1].split('=')[1];
          var name = params[2].split('=')[1];

          session.dialogData.username = name;
          if(token && code){
            session.beginDialog('validateCode', { code : code });

          } else {
            session.replaceDialog('signinPrompt', { invalid: true });
          }
      } else {
          session.replaceDialog('signinPrompt', { invalid: true });
      }
    } else {
        session.replaceDialog('signinPrompt', { invalid: true });
    }
  },
  function (session, results) {
      if (results.response) {
          //code validated
          session.endDialogWithResult({ response: true, username: session.dialogData.username});
      }else {
          session.endDialogWithResult({ response: false});
      }
  }
]);

bot.dialog('validateCode', [
  function (session, args) {
    if (!session.dialogData.code && args && args.code) {
      session.dialogData.code = args.code;
    }
    builder.Prompts.text(session, "Please enter the code on the screen");
  },
  function (session, results) {
    var code = results.response;
    if (code === session.dialogData.code) {
        // // Store refresh token
        // session.userData.refreshToken = session.dialogData.token;
        // session.sendTyping();
        // // ... Async call to convert refresh token to access token
        // var accessToken = '';
        // session.privateConversationData.accessToken = accessToken;
        session.endDialogWithResult({ response: true });
    } else {
        session.send("hmm... Looks like that was an invalid code. Please try again.");
        session.replaceDialog('validateCode', {code: session.dialogData.code});
    }
  }
]);