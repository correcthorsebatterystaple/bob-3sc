let client;
let access_token;

function initClient() {
  client = google.accounts.oauth2.initTokenClient({
    client_id:
      "106113587711-4v9qoovue479a15nva7masil8n92v4uv.apps.googleusercontent.com",
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly",
    callback: (tokenResponse) => {
      access_token = tokenResponse.access_token;
      document.getElementById("login-btn").disabled = true;
      document.getElementById("fetch-btn").disabled = false;
      document.getElementById("logout-btn").disabled = false;
    },
  });
}
function getToken() {
  client.requestAccessToken();
}
function revokeToken() {
  google.accounts.oauth2.revoke(access_token, () => {
    console.log("access token revoked");
    access_token = null;
    document.getElementById("login-btn").disabled = false;
    document.getElementById("fetch-btn").disabled = true;
    document.getElementById("logout-btn").disabled = true;
  });
}
