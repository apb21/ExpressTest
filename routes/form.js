var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

router.get('/', async function(req,res,next){
  let parms = { title : 'Form' };

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;

    // Initialize Graph client
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
    res.render('form',parms);
  } else {
    res.redirect('/');
  }
});

router.get('/:record', async function(req,res,next){
  let parms = { title : 'Form' };
  Object.assign(parms, req.params, parms);

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;

    // Initialize Graph client
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    // TODO: Load data from OneDrive for requested record (if it exists, otherwise load blank)

    console.log(parms);
    res.render('form',parms);
  }
  else{
    res.redirect('/');
  };
});

router.post('/', async function(req,res,next){

  // TODO: Save data to OneDrive, in same file (if exists) or create new file.
  console.log(req.body);
  res.redirect('/');
});

module.exports = router;
