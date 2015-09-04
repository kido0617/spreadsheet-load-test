using Google.GData.Client;
using Google.GData.Spreadsheets;
using System;
using System.Collections.Generic;

namespace SpreadSheetLoader {
  class SpreadSheetManager {

    SpreadsheetsService service;
    SpreadsheetEntry spreadSheetEntry;

    public SpreadSheetManager(OAuth2Parameters parameters) {
      GOAuth2RequestFactory requestFactory =
      new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
      service = new SpreadsheetsService("MySpreadsheetIntegration-v1");
      service.RequestFactory = requestFactory;
    }

    public void LoadSpreadSheet(string spreadSheetId) {
      SpreadsheetQuery query = new SpreadsheetQuery();
      query.Uri = new Uri("https://spreadsheets.google.com/feeds/spreadsheets/private/full/" + spreadSheetId);

      SpreadsheetFeed feed = service.Query(query);
      spreadSheetEntry = (SpreadsheetEntry)feed.Entries[0];
    }

    public ListFeed GetListFeed(int sheetNum) {
      WorksheetFeed wsFeed = spreadSheetEntry.Worksheets;
      WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[sheetNum];

      AtomLink listFeedLink = worksheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);

      ListQuery listQuery = new ListQuery(listFeedLink.HRef.ToString());
      return service.Query(listQuery);
    }

    public CellFeed GetCellFeed(int sheetNum) {
      WorksheetFeed wsFeed = spreadSheetEntry.Worksheets;
      WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[sheetNum];
      CellQuery cellQuery = new CellQuery(worksheet.CellFeedLink);
      CellFeed cellFeed = service.Query(cellQuery);
      return cellFeed;

    }


    public List<string> GetFirstRow(int sheetNum) {
      WorksheetFeed wsFeed = spreadSheetEntry.Worksheets;
      WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[sheetNum];
      CellQuery cellQuery = new CellQuery(worksheet.CellFeedLink);
      cellQuery.MinimumRow = 1;
      cellQuery.MaximumRow = 1;

      CellFeed cellFeed = service.Query(cellQuery);

      List<string> list = new List<string>();

      foreach (CellEntry cell in cellFeed.Entries) {
        list.Add(cell.Value);
      }
      return list;
    }


  }
}
