var express = require('express');
var router = express.Router();

/* GET users listing. */
router.get('/', function(req, res, next) {
  res.render('index',{ title: 'Users', active: {'home':false,'users':true} });
});
router.get('/:user_id',function(req,res){
  res.send(req.params['user_id']);
})

module.exports = router;
