var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

/* GET users OneDrive */
router.get('/', async function(req, res, next){
  let parms = { title:'Drive' };

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName){
    parms.user = userName;

    // Initialize Graph client
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    try{
      const result = await client
      .api('me/drive/special/approot/children')
      .get();
      //console.log(result);
      parms.drive = result;
      parms.ItemCount = result.value.length;
    }
    catch (err){
      parms.error = { status: `${err.code}: ${err.message}` };
      res.render('error',parms);
    }
  }
  else{
    //No User so Redirect Home
    res.redirect('/');
  };
  res.render('drive',parms);
});
module.exports = router;
