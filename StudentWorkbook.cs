/// StudentWorkbook
/// This class is very specific to Excel spreadsheets written for
/// TCES 481, in-class reviews of student presentations. This class
/// opens a particular workbook for a particular student, extracts 
/// the Team Number of the group being reviewed, extracts the scores
/// for the various categories, and extracts the Additional Comments.
/// 
/// R. Gutmann
/// 1/23/2016
/// 


using System;
using Excel = Microsoft.Office.Interop.Excel;


namespace ClassAssessments
{
  class StudentWorkbook
  {
    Excel.Workbook excelWorkbook;

    int teamNumber = 0;      // team number being reviewed
    public int TeamNumber
    {
      get { return teamNumber; }
    }

    Excel.Range scores;  // copy of the scores 
    public Excel.Range Scores
    {
      get { return scores; }
    }

    Excel.Range comments;  // copy of the scores 
    public Excel.Range Comments
    {
      get { return comments; }
    }

    /// <summary>
    /// Class constructor does all the work
    /// </summary>
    /// <param name="excelApp"></param>
    /// <param name="fileName"></param>
    public StudentWorkbook( Excel.Application excelApp, string fileName )
    {
      try  // to open the student's spreadsheet
      {
        excelWorkbook = excelApp.Workbooks.Open( fileName, 0,
            true, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true,  // read only
            false, 0, true, false, false );
      }
      catch ( Exception e )
      {
        Console.WriteLine( "error: " + e.Message );
        Console.WriteLine( "Could not open spreadsheet " + fileName );
      }

      Excel.Sheets excelSheets = excelWorkbook.Worksheets;  // get the Worksheets collection
      Excel.Worksheet excelWorksheet = excelSheets[ 1 ];    // get the first one

      // get the Team Number cell
      Excel.Range excelCell = (Excel.Range)excelWorksheet.get_Range( "B4", "B4" );

      // try to convert this cell to an integer
      if ( ( teamNumber = TryForInt( excelCell.Value ) ) == 0 )
      {
        Console.WriteLine( "\nTeam number invalid in " + fileName + "\n" );
      }

      // get the scores cells
      scores = excelWorksheet.get_Range( "B7", "B15" );

      // get the Additional Comments cell
      comments = excelWorksheet.get_Range( "B18", "B18" );

    }  // end of StudentWorkbook()


    /// <summary>
    /// Close the workbook
    /// </summary>
    public void Close()
    {
      excelWorkbook.Close( false );  // don't save 'changes'
    }  // end Close()


    /// <summary>
    /// Try to convert a cell to an int. 
    /// Cell may be blank (null) or
    /// might be a string or 
    /// might be a number (double)
    /// Return a 0 if conversion cannot be made
    /// </summary>
    /// <param name="o"></param>
    /// <returns></returns>
    int TryForInt( object o )
    {
      int retval = 0;
      if ( o != null )
      {
        if ( o.GetType() == typeof( double ) )
        {
          retval = (int)( (double)o );
        }
        else if ( o.GetType() == typeof( string ) )
        {
          double d = 0;
          double.TryParse( (string)o, out d );
          retval = (int)d;
        }
      }
      return retval;
    }

  }  // end class StudentWorkbook
}  // end namespace
