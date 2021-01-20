/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file is the main Node.js server file that defines the express middleware.
 */

if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}
var createError = require('http-errors');
var express = require('express');
var path = require('path');
var fetch = require('node-fetch');
var form = require('form-urlencoded').default;
var graphModule = require('./src/helpers/msgraph-helper');

var authRouter = require('./src/authRoute');
const { debug } = require('webpack');
var Promise = require('promise');

var app = express();


app.use(express.json());
app.use(express.urlencoded({ extended: false }));


/* Turn off caching when developing */
if (process.env.NODE_ENV !== 'production') {
  app.use(express.static(path.join(__dirname, 'public')));

  app.use(function (req, res, next) {
    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    res.header('Acces-Control-Allow-Origin','*');
    res.header('Acces-Control-Allow-Methods','GET,POST,PUT,PATCH,DELETE');
    res.header('Acces-Contorl-Allow-Methods','Content-Type','Authorization');
    next()
  });
} else {
  // In production mode, let static files be cached.
 app.use(express.static(path.join(__dirname, 'public')));
}


app.get('/index.html', (async (req, res) => {
  res.sendFile('src/index.html', {root: __dirname });
}));

app.get('/home/commands',function(req,res) {
  res.sendFile('src/commands/commands.html', {root: __dirname });
});

app.get('/taskpane.html', (async (req, res) => {
  res.sendFile('src/taskpane/taskpane.html', {root: __dirname });
}));

app.get('/public/taskpane.css', (async (req, res) => {
  res.sendFile('public/taskpane.css', {root: __dirname });
}));

app.get('/public/taskpane.js', (async (req, res) => {
  res.sendFile('public/taskpane.js', {root: __dirname });
}));

app.get('/public/index.js', (async (req, res) => {
  res.sendFile('public/index.js', {root: __dirname });
}));


app.get('/auth', async function(req, res, next) {
  const authorization = req.get('Authorization');
  let itemId = req.get('Id');
  const search = '_';
  const replacer = new RegExp(search, 'g');
  const search1 = '/';
  const replacer1 = new RegExp(search1, 'g');
  itemId = itemId.replace(replacer, '+');
  const restId = itemId.replace(replacer1, '-')
  console.log('auth');
  if (authorization == null) {
     let error = new Error('No Authorization header was found.');
     next(error);
  } 
  else {
    console.log('else');
    const [schema, jwt] = authorization.split(' ');
    const formParams = {
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt,
      requested_token_use: 'on_behalf_of',
      scope: ['Files.Read.All'].join(' ')
    };

    const stsDomain = 'https://login.microsoftonline.com';
    const tenant = process.env.TENANT_ID;
    const tokenURLSegment = 'oauth2/v2.0/token';

    try {
      const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
        method: 'POST',
        body: form(formParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
      });
      const json = await tokenResponse.json();
      console.log(`access token:=> ${json.access_token}` );

     
      const graphApiCallRes = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}`, {
        method: 'DELETE',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Authorization': 'Bearer ' + json.access_token,
          'Cache-Control': 'private, no-cache, no-store, must-revalidate',
          'Expires': '-1',
          'Pragma': 'no-cache'
        }
      });
      const json1 = await graphApiCallRes.json();
      console.log('json1 ' + JSON.stringify(json1));
      res.send(json1);
    }
    catch(error) {
      console.log('error' + error);
      res.status(500).send(error);
    }
  }
});


module.exports = app;
