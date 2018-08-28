var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

router.get('/', async function(req,res,next){
  let parms = { title : 'Form', app_name: process.env.APP_NAME };
  parms.data = ["","","","","","","","","",""];
  parms.action = '/form';

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;
    res.render('form',parms);
  } else {
    res.redirect('/');
  }
});

router.get('/:record', async function(req,res,next){
  let parms = { title : 'Form' };
  Object.assign(parms, req.params, parms);
  parms.action = '/form/'+parms.record;

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

    // TODO: Load data from Excel workbook on OneDrive for requested record (if it exists, otherwise load blank)
    try{
      const pathTemplate = 'me/drive/root:/Apps/APP_NAME/APP_NAME.xlsx:/workbook/tables(%271%27)/rows/itemAt(index=ROW)';//?$top=1&$skip=ROW';
      var path = pathTemplate.replace("APP_NAME",process.env.APP_NAME);
      var pathFinder = pathTemplate.search('APP_NAME');
      while(pathFinder > -1){
          var path = path.replace("APP_NAME",process.env.APP_NAME);
          var pathFinder = path.search('APP_NAME');
      }
      var path = path.replace('ROW',parms.record);
      //console.log(path);
      const result = await client
        .api(path)
        .get();
        // TODO: Fill In Form
        parms.data = result.values[0];
    } catch(err){
      parms.error = { status: `${err.code}: ${err.message}` };
      res.render('error',parms);
    }
    //console.log(parms);
    res.render('form',parms);
  }
  else{
    res.redirect('/');
  };
});

router.post('/', async function(req,res,next){
  let parms = { 'title' : 'Form'};
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
    let valuesArray = [];
    for (var i in req.body){
      valuesArray.push("'"+req.body[i].toString());
    }

    const config = {
      'index' : null,
      'values': [valuesArray]
    };
    //console.log(config);
    try{
      const pathTemplate = 'me/drive/root:/Apps/APP_NAME/APP_NAME.xlsx:/workbook/tables(%271%27)/rows/add';
      var path = pathTemplate.replace("APP_NAME",process.env.APP_NAME);
      var pathFinder = pathTemplate.search('APP_NAME');
      while(pathFinder > -1){
          var path = path.replace("APP_NAME",process.env.APP_NAME);
          var pathFinder = path.search('APP_NAME');
      }
      const result = await client
        .api(path)
        .post(config)
        .then((res) =>{
            // TODO: Respond with outcome to user
          })
        .catch((err) =>{
          parms.error = { status: `${err.code}: ${err.message}` };
          res.render('error',parms);
        });

    } catch(err){
      parms.error = { status: `${err.code}: ${err.message}` };
      res.render('error',parms);
    }
  }else{
    res.redirect('/');
  };

  res.redirect('/drive');
});

router.post('/:record', async function(req,res,next){
  let parms = { 'title' : 'Form'};
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

    let valuesArray = [];

    for (var i in req.body){
      valuesArray.push("'"+req.body[i].toString());
    }

    const config = {
      'values': [valuesArray]
    };
    //console.log(config);
    try{
      const pathTemplate = 'me/drive/root:/Apps/APP_NAME/APP_NAME.xlsx:/workbook/tables(%271%27)/rows/$/itemAt(index=ROW)';
      var path = pathTemplate.replace("APP_NAME",process.env.APP_NAME);
      var pathFinder = pathTemplate.search('APP_NAME');
      while(pathFinder > -1){
          var path = path.replace("APP_NAME",process.env.APP_NAME);
          var pathFinder = path.search('APP_NAME');
      }
      var path = path.replace("ROW",parms.record);
      const result = await client
        .api(path)
        .patch(config)
        .then((res) =>{
            // TODO: Respond with outcome to user?
          })
        .catch((err) =>{
          parms.error = { status: `${err.code}: ${err.message}` };
          res.render('error',parms);
        });

    } catch(err){
      parms.error = { status: `${err.code}: ${err.message}` };
      res.render('error',parms);
    }
  }else{
    res.redirect('/');
  };

  res.redirect('/drive');
});

module.exports = router;
