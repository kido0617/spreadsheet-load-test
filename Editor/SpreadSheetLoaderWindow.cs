using System.Net;
using UnityEditor;
using UnityEngine;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Google.GData.Spreadsheets;

namespace SpreadSheetLoader {

  public class SpreadSheetLoaderWindow : EditorWindow {

    OAuth2 oAuth2 = new OAuth2();
    SpreadSheetManager spreadSheet;

    string accessCode = "";
    readonly string PREF_STR = "refreshToken";

    [MenuItem("Window/SpreadSheetLoader")]
    static void Open() {
      EditorWindow.GetWindow<SpreadSheetLoaderWindow>("SpreadSheetLoader");
      ServicePointManager.ServerCertificateValidationCallback = Validator;
    }

    public static bool Validator(object in_sender, X509Certificate in_certificate, X509Chain in_chain, SslPolicyErrors in_sslPolicyErrors) {
      return true;
    }

    void OnGUI() {
      if (GUILayout.Button("Get Access Code")) {
        string authUrl = oAuth2.GetAuthURL();
        Application.OpenURL(authUrl);
      }
      accessCode = EditorGUILayout.TextField("Access Code", accessCode);
      if (GUILayout.Button("Authentication with Acceess Code")) {
        string refreshToken = oAuth2.AuthWithAccessCode(accessCode);
        EditorPrefs.SetString(PREF_STR, refreshToken);
      }

      if (GUILayout.Button("Load Spread Sheet")) {
        if (EditorPrefs.GetString(PREF_STR) == "") {
          Debug.LogError("Refresh Token is not set. You need above authentication steps.");
          return;
        }
        Load();
      }
    }

    void Load() {
      var refreshToken = EditorPrefs.GetString(PREF_STR);
      spreadSheet = new SpreadSheetManager(oAuth2.GetOAuth2Parameter(refreshToken));
      spreadSheet.LoadSpreadSheet(Config.SPREAD_SHEET_ID);
      var listFeed = spreadSheet.GetListFeed(0);      //set sheet number

      //Display results
      foreach (ListEntry row in listFeed.Entries) {
        foreach (ListEntry.Custom element in row.Elements) {
          Debug.Log(element.Value);
        }
      }

    }

  }
}