using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassAssessments
{
  class TeamWorkbook
  {
    Excel.Application excelApp;
    //string fileName;
    Excel.Workbook excelWorkbook = null;
    Excel.Worksheet excelWorksheet = null;
    int currentRow = 6;


    public TeamWorkbook( Excel.Application app, string fileName )
    {
      excelApp = app;
      //fileName = fn;

      //DirectoryInfo topDir = new DirectoryInfo( Environment.CurrentDirectory );
      //string dirName = topDir.FullName;
      //Console.WriteLine( dirName );
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
      scores.Copy();
      string newCell = "A" + (++currentRow).ToString();  // increment row
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
      string newCell = "J" + ( currentRow ).ToString();
      Excel.Range excelRange = excelWorksheet.get_Range( newCell, newCell );  // starts with A7, then moves down

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
