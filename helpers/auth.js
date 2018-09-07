const jwt = require('jsonwebtoken');

let newAgent = null;
const proxy = process.env.HTTP_PROXY || process.env.http_proxy || null;
if (proxy !== null) {
  const HttpProxyAgent = require('http-proxy-agent');
  newAgent = new HttpProxyAgent(proxy);
}

const credentials = {
  client: {
    id: process.env.APP_ID,
    secret: process.env.APP_PASSWORD,
  },
  auth: {
    tokenHost: 'https://login.microsoftonline.com',
    authorizePath: 'common/oauth2/v2.0/authorize',
    tokenPath: 'common/oauth2/v2.0/token',
  },
  http: {
    agent: newAgent,
  },
};
const oauth2 = require('simple-oauth2').create(credentials);

function getAuthUrl() {
  const returnVal = oauth2.authorizationCode.authorizeURL({
    redirect_uri: process.env.REDIRECT_URI,
    scope: process.env.APP_SCOPES,
  });
  return returnVal;
}

exports.getAuthUrl = getAuthUrl;

async function getTokenFromCode(authCode, res) {
  const tokenConfig = {
    code: authCode,
    redirect_uri: process.env.REDIRECT_URI,
    scope: process.env.APP_SCOPES,
  };
  try {
    const result = await oauth2.authorizationCode.getToken(tokenConfig);
    const token = oauth2.accessToken.create(result);
    saveValuesToCookie(token, res);
    return token.token.access_token;
  } catch (err) {
    const parms = {};
    parms.error = { status: `${err.code}: ${err.message}` };
    res.render('error', parms);
  }
}

exports.getTokenFromCode = getTokenFromCode;

async function saveValuesToCookie(token, res) {
  // Parse the identity token
  const user = jwt.decode(token.token.id_token);

  // Save the access token in a cookie
  res.cookie('graph_access_token', token.token.access_token, { maxAge: 3600000, httpOnly: true });
  // Save the user's name in a cookie
  res.cookie('graph_user_name', user.name, { maxAge: 3600000, httpOnly: true });
  // Save the refresh token in a cookie
  res.cookie('graph_refresh_token', token.token.refresh_token, { maxAge: 7200000, httpOnly: true });
  // Save the token expiration time in a cookie
  res.cookie('graph_token_expires', token.token.expires_at.getTime(), { maxAge: 3600000, httpOnly: true });
}

async function clearCookies(res) {
  // Clear cookies
  res.clearCookie('graph_access_token', { maxAge: 3600000, httpOnly: true });
  res.clearCookie('graph_user_name', { maxAge: 3600000, httpOnly: true });
  res.clearCookie('graph_refresh_token', { maxAge: 7200000, httpOnly: true });
  res.clearCookie('graph_token_expires', { maxAge: 3600000, httpOnly: true });
}

exports.clearCookies = clearCookies;

async function getAccessToken(cookies, res) {
  try {
    // Do we have an access token cached?
    const token = cookies.graph_access_token;

    if (token) {
      // We have a token, but is it expired?
      // Expire 5 minutes early to account for clock differences
      const FIVE_MINUTES = 300000;
      const expiration = new Date(parseFloat(cookies.graph_token_expires - FIVE_MINUTES));
      if (expiration > new Date()) {
        // Token is still good, just return it
        return token;
      }
    }

    // Either no token or it's expired, do we have a
    // refresh token?
    const refreshToken = cookies.graph_refresh_token;
    if (refreshToken) {
      const newToken = await oauth2.accessToken.create({ refreshToken }).refresh();
      saveValuesToCookie(newToken, res);
      return newToken.token.access_token;
    }
  } catch (err) {
    // Nothing in the cookies that helps, return empty
    return null;
  }
}

exports.getAccessToken = getAccessToken;
