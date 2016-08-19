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

// server.get('/login',
//   function (req, res, next) {
//     console.log('Login req');
//     passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', state: req.query.state });
//   },
//   function (req, res) {
//     console.log('Login');
//   });

server.get('/api/OAuthCallback/',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', state: 'ritatest' }),
  function (req, res) {
    console.log('Returned from AzureAD.');
    console.log('state: ' + req.query.state);
    var botdata = req.query.state;
    console.log('botdata: ' + botdata);
    var addressId = botdata.split('|')[0];
    var conversationId = botdata.split('|')[1];
    console.log('addressId: ' + addressId + '|' + 'conversationId: ' + conversationId);

    var msg =  { type: 'message',
  timestamp: '2016-08-19T23:59:21.8521273Z',
  text: 'hi',
  attachments: [],
  entities: [],
  address: 
   { id: addressId,
     channelId: 'webchat',
     user: { id: 'FVCBFTKEVD2', name: 'FVCBFTKEVD2' },
     conversation: { id: conversationId },
     bot: { id: 'authbot', name: 'authbot' },
     serviceUrl: 'https://webchat.botframework.com',
     useAuth: true },
  source: 'webchat',
  agent: 'botbuilder',
  user: { id: 'FVCBFTKEVD2', name: 'FVCBFTKEVD2' } };

  var address = { id: addressId,
     channelId: 'webchat',
     user: { id: 'FVCBFTKEVD2', name: 'FVCBFTKEVD2' },
     conversation: { id: conversationId },
     bot: { id: 'authbot', name: 'authbot' },
     serviceUrl: 'https://webchat.botframework.com',
     useAuth: true };

    console.log('before bot receive');
    bot.receive(msg); 
    console.log('after bot receive');
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
  	//console.log(profile);
    //console.log('passport.use: ' + req);

    if (!profile.email) {
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

bot.dialog('/', function (session) {
    console.log('bot dialog');
    console.log(session.message);
    if (!session.userData.users) {
        session.userData.users = [];
    } else {

    }
    console.log("message: " + session.message);
    console.log("address: " + session.message.address);
    console.log("convo: " + session.message.address.conversation.id);
    //var state = querystring.stringify(session.message.address);
    var state = session.message.address.id + "|" + session.message.address.conversation.id;

    console.log("state: " + state);
    session.send("Hi there! Welcome! Please click %s to sign in: ", "https://authbot.azurewebsites.net/login?state=" + state);
    
});