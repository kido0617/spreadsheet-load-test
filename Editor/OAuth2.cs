using Google.GData.Client;

namespace SpreadSheetLoader {
  class OAuth2 {

    string SCOPE = "https://spreadsheets.google.com/feeds";
    string REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob";

    public string GetAuthURL() {
      OAuth2Parameters parameters = new OAuth2Parameters();
      parameters.ClientId = Config.CLIENT_ID;
      parameters.ClientSecret = Config.CLIENT_SECRET;
      parameters.RedirectUri = REDIRECT_URI;
      parameters.Scope = SCOPE;
      return OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
    }

    public string AuthWithAccessCode(string accessCode) {
      OAuth2Parameters parameters = new OAuth2Parameters();
      parameters.ClientId = Config.CLIENT_ID;
      parameters.ClientSecret = Config.CLIENT_SECRET;
      parameters.RedirectUri = REDIRECT_URI;
      parameters.Scope = SCOPE;
      parameters.AccessCode = accessCode;

      OAuthUtil.GetAccessToken(parameters);
      return parameters.RefreshToken;

    }

    public OAuth2Parameters GetOAuth2Parameter(string refreshToken) {
      OAuth2Parameters parameters = new OAuth2Parameters();
      parameters.ClientId = Config.CLIENT_ID;
      parameters.ClientSecret = Config.CLIENT_SECRET;
      parameters.RefreshToken = refreshToken;

      OAuthUtil.RefreshAccessToken(parameters);
      return parameters;

    }
  }
}
