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
    parms.app_name = process.env.APP_NAME;

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
      for (var i in result.value){
        //console.log(result.value[i].name);
        if (result.value[i].name == process.env.APP_NAME+'.xlsx'){
          try{
            const pathTemplate = 'me/drive/root:/Apps/APP_NAME/APP_NAME.xlsx:/workbook/tables(%271%27)/rows';
            var path = pathTemplate.replace("APP_NAME",process.env.APP_NAME);
            var pathFinder = pathTemplate.search('APP_NAME');
            while(pathFinder > -1){
                var path = path.replace("APP_NAME",process.env.APP_NAME);
                var pathFinder = path.search('APP_NAME');
            }
            const resultlist = await client
            .api(path)
            .get();
            //console.log(resultlist.value);
            parms.resultlist = resultlist;
            parms.resultlistcount = resultlist.value.length;
          }catch(err){
            parms.error = { status: `${err.code}: ${err.message}` };
            res.render('error',parms);
          }
        }
      }
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
