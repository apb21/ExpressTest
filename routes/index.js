const express = require('express');

const router = express.Router();
const authHelper = require('../helpers/auth');
const graph = require('@microsoft/microsoft-graph-client');

/* GET home page. */
router.get('/', async (req, res, next) => {
  const parms = { title: 'Home', app_name: process.env.APP_NAME };

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;

    // Initialize Graph client
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    try {
      const result = await client
        .api('me/photos/48x48/$value')
        .get();
      parms.photoBlob = result;
    } catch (err) {
      parms.error = { status: `${err.code}: ${err.message}` };
      res.render('error', parms);
    }
  } else {
    parms.signInUrl = authHelper.getAuthUrl();
  }
  res.render('index', parms);
});

module.exports = router;
