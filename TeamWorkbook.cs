/// TeamWorkbook
/// 
/// This class is very specific to Excel spreadsheets written for
/// TCES 481, in-class reviews of student presentations. This class
/// opens a particular workbook for a particular team and provides 
/// utility methods to paste scores and paste comments from individual
/// student assessments.
/// 
/// R. Gutmann
/// 1/23/2016
/// 



using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClassAssessments
{
  class TeamWorkbook
  {
    Excel.Workbook    excelWorkbook = null;
    Excel.Worksheet   excelWorksheet = null;
    int currentRow = 6;  // which row to paste scores and comments 


    public TeamWorkbook( Excel.Application excelApp, string fileName )
    {
      try
      {
        excelWorkbook = excelApp.Workbooks.Open( fileName, 0,
            false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true,
            false, 0, true, false, false );
      }
      catch (Exception e)
      {
        Console.WriteLine( e.Message );
        Console.WriteLine( "Could not open spreadsheet " + fileName );
      }

      // The following gets the Worksheets collection
      Excel.Sheets excelSheets = excelWorkbook.Worksheets;

      string currentSheet = "Sheet1";  // get Sheet1 to operate on
      excelWorksheet = (Excel.Worksheet)excelSheets.get_Item( currentSheet );
    }  // end TeamWorkbook()



    public void PasteScores( Excel.Range scores )
    {
      // note: for PasteSpecial to work, both the student and team workbooks must be opend using the same Excel App
      scores.Copy(); // copy scoares to clipboard
      string newCell = "A" + (++currentRow).ToString();  // Note: increment row
      Excel.Range excelRange = excelWorksheet.get_Range( newCell, newCell );  // starts with A7, then moves down

      try {
        excelRange.PasteSpecial( Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                                 Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                 false,
                                 true );  // transpose the column into a row
      }
      catch( Exception e )
      {
        Console.WriteLine( e.Message );
      }
    }  // end PasteScores()


    public void PasteComments( Excel.Range comments )
    {
      comments.Copy();
      string newCell = "J" + ( currentRow ).ToString();   // very specific
      Excel.Range excelRange = excelWorksheet.get_Range( newCell, newCell );  

      try
      {
        excelRange.PasteSpecial( Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                                 Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                 false,
                                 false );  
      }
      catch ( Exception e )
      {
        Console.WriteLine( e.Message );
      }

    }  // end PasteComments()


    public void Close()
    {
      excelWorkbook.Close( true );  // save changes
    }  // end Close()



    static public int GetTeamNumber( string team )
    {
      int teamNumber;
      if ( !int.TryParse( team, out teamNumber ) )
      {
        throw new ArgumentException( "Not a valid team number", "teamNumber" );
      }
      return teamNumber;
    }  // end GetTeamNumber()



    static public string GetTeamPath( string path, string team )
    {
      string teamString = "Team" + team;  // example: Team01
      string retval = path + "\\" + teamString + "\\" + teamString + "P1.xlsx";
      return retval;
    }  // end GetTeamPath()


  }  // end class TeamWorkbook
}  // end namespace
