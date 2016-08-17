var restify = require('restify');
var builder = require('botbuilder');
var passport = require('passport');
var OIDCStrategy = require('passport-azure-ad').OIDCStrategy;

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
server.use(passport.initialize());

server.get('/login',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login' }),
  function (req, res) {
    console.log('Login');
  });

server.get('/api/OAuthCallback',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login' }),
  function (req, res) {
    console.log('Returned from AzureAD.');
    console.log(req);
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

passport.use(new OIDCStrategy({
    callbackURL: 'https://authbot.azurewebsites.net/api/OAuthCallback',
    realm: 'common',
    clientID: process.env.MICROSOFT_APP_ID,
    clientSecret: process.env.MICROSOFT_APP_PASSWORD,
    identityMetadata: 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
    skipUserProfile: true,
    responseType: 'code',
    responseMode: 'query',
    scope: ['email', 'profile']
  },
  function(iss, sub, profile, accessToken, refreshToken, done) {
  	console.log(profile);

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
          users.push(profile);
          return done(null, profile);
        }
        return done(null, user);
      });
    });
  }
));
//=========================================================
// Bots Dialogs
//=========================================================

bot.dialog('/', function (session) {
    session.send("Hello World");
});