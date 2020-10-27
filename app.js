const express = require('express');
const app = express();

const bodyParser = require('body-parser');
const cookieParser = require('cookie-parser');
const session = require('express-session');
const moment = require('moment');
const querystring = require('querystring');
const outlook = require('node-outlook');
const port = 3000;

const pages = require('./pages');
const authHelper = require('./authHelper');
const authHelp = new authHelper();

class Server {
  constructor() {
    this.ExpressMiddleware();
    this.Routes();
    this.Start();
  }

  ExpressMiddleware() {
    app.use(express.static('static'));
    app.use(bodyParser.json());
    app.use(cookieParser());
    app.use(session(
      { secret: '0dc529ba-5051-4cd6-8b67-c9a901bb8bdf',
        resave: false,
        saveUninitialized: false 
      }));
  }
  Routes() {
    app.get('/', function(req, res) {
      res.send(pages.loginPage(authHelp.getAuthUrl()));
    });
    
    app.get('/authorize', function(req, res) {
      const authCode = req.query.code;
      if (authCode) {
        console.log('');
        console.log('Retrieved auth code in /authorize: ' + authCode);
        authHelp.getTokenFromCode(authCode, tokenReceived, req, res);
      }
      else {
        // redirect to home
        console.log('/authorize called without a code parameter, redirecting to login');
        res.redirect('/');
      }
    });
    
    app.get('/logincomplete', function(req, res) {
      const access_token = req.session.access_token;
      const refresh_token = req.session.access_token;
      const email = req.session.email;
      
      if (access_token === undefined || refresh_token === undefined) {
        console.log('/logincomplete called while not logged in');
        res.redirect('/');
        return;
      }
      
      res.send(pages.loginCompletePage(email));
    });
    
    app.get('/refreshtokens', function(req, res) {
      const refresh_token = req.session.refresh_token;
      if (refresh_token === undefined) {
        console.log('no refresh token in session');
        res.redirect('/');
      }
      else {
        authHelp.getTokenFromRefreshToken(refresh_token, tokenReceived, req, res);
      }
    });
    
    app.get('/logout', function(req, res) {
      req.session.destroy();
      res.redirect('/');
    });
    
    app.get('/sync', function(req, res) {
      const token = req.session.access_token;
      const email = req.session.email;
      if (token === undefined || email === undefined) {
        console.log('/sync called while not logged in');
        res.redirect('/');
        return;
      }
      
      // Set the endpoint to API v2
      outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
      // Set the user's email as the anchor mailbox
      outlook.base.setAnchorMailbox(req.session.email);
      // Set the preferred time zone
      outlook.base.setPreferredTimeZone('Eastern Standard Time');
      
      // Use the syncUrl if available
      const requestUrl = req.session.syncUrl;
      if (requestUrl === undefined) {
        // Calendar sync works on the CalendarView endpoint
        requestUrl = outlook.base.apiEndpoint() + '/Me/CalendarView';
      }
      
      // Set up our sync window from midnight on the current day to
      // midnight 7 days from now.
      const startDate = moment().startOf('day');
      const endDate = moment(startDate).add(7, 'days');
      // The start and end date are passed as query parameters
      const params = {
        startDateTime: startDate.toISOString(),
        endDateTime: endDate.toISOString()
      };
      
      // Set the required headers for sync
      const headers = {
        Prefer: [ 
          // Enables sync functionality
          'odata.track-changes',
          // Requests only 5 changes per response
          'odata.maxpagesize=5'
        ]
      };
      
      const apiOptions = {
        url: requestUrl,
        token: token,
        headers: headers,
        query: params
      };
      
      outlook.base.makeApiCall(apiOptions, function(error, response) {
        if (error) {
          console.log(JSON.stringify(error));
          res.send(JSON.stringify(error));
        }
        else {
          if (response.statusCode !== 200) {
            console.log('API Call returned ' + response.statusCode);
            res.send('API Call returned ' + response.statusCode);
          }
          else {
            const nextLink = response.body['@odata.nextLink'];
            if (nextLink !== undefined) {
              req.session.syncUrl = nextLink;
            }
            const deltaLink = response.body['@odata.deltaLink'];
            if (deltaLink !== undefined) {
              req.session.syncUrl = deltaLink;
            }
            res.send(pages.syncPage(email, response.body.value));
          }
        }
      });
    });
    
    app.get('/viewitem', function(req, res) {
      const itemId = req.query.id;
      const access_token = req.session.access_token;
      const email = req.session.email;
      
      if (itemId === undefined || access_token === undefined) {
        res.redirect('/');
        return;
      }
      
      const select = {
        '$select': 'Subject,Attendees,Location,Start,End,IsReminderOn,ReminderMinutesBeforeStart'
      };
      
      const getEventParameters = {
        token: access_token,
        eventId: itemId,
        odataParams: select
      };
      
      outlook.calendar.getEvent(getEventParameters, function(error, event) {
        if (error) {
          console.log(error);
          res.send(error);
        }
        else {
          res.send(pages.itemDetailPage(email, event));
        }
      });
    });
    
    app.get('/updateitem', function(req, res) {
      const itemId = req.query.eventId;
      const access_token = req.session.access_token;
      
      if (itemId === undefined || access_token === undefined) {
        res.redirect('/');
        return;
      }
      
      const newSubject = req.query.subject;
      const newLocation = req.query.location;
      
      console.log('UPDATED SUBJECT: ', newSubject);
      console.log('UPDATED LOCATION: ', newLocation);
      
      const updatePayload = {
        Subject: newSubject,
        Location: {
          DisplayName: newLocation
        }
      };
      
      const updateEventParameters = {
        token: access_token,
        eventId: itemId,
        update: updatePayload
      };
      
      outlook.calendar.updateEvent(updateEventParameters, function(error, event) {
        if (error) {
          console.log(error);
          res.send(error);
        }
        else {
          res.redirect('/viewitem?' + querystring.stringify({ id: itemId }));
        }
      });
    });
    
    app.get('/deleteitem', function(req, res) {
      const itemId = req.query.id;
      const access_token = req.session.access_token;
      
      if (itemId === undefined || access_token === undefined) {
        res.redirect('/');
        return;
      }
      
      const deleteEventParameters = {
        token: access_token,
        eventId: itemId
      };
      
      outlook.calendar.deleteEvent(deleteEventParameters, function(error, event) {
        if (error) {
          console.log(error);
          res.send(error);
        }
        else {
          res.redirect('/sync');
        }
      });
    });
    
  }
    
  Start() {
    app.listen(port, () => console.log(`server is listenning on ${port}`));
  }

}

new Server();