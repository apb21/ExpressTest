const express = require('express');

const router = express.Router();
const authHelper = require('../helpers/auth');

/* GET /authorize. */
router.get('/', async (req, res, next) => {
  // Get auth code
  const code = req.query.code;

  // If code is present, use it
  if (code) {
    // let token

    try {
      // token =
      await authHelper.getTokenFromCode(code, res);
    } catch (error) {
      res.render('error', { title: 'Error', message: 'Error exchanging code for token', error });
    }
  } else {
    // Otherwise complain
    res.render('error', { title: 'Error', message: 'Authorization error', error: { status: 'Missing code parameter' } });
  }

  // Redirect to home
  res.redirect('/');
});

/* GET /authorize/signout */
router.get('/signout', (req, res, next) => {
  authHelper.clearCookies(res);

  // Redirect to home
  res.redirect('/');
});

module.exports = router;
