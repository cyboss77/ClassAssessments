///
/// Class Assessments
/// This program is a utility program specific to TCES 481 (Senior Projects)
/// Students evaluate team presentations and turn in Excel spreadsheets with
/// scores for various aspects of the presentation. This program builds a 
/// summary spreadsheet for each team. It copies and pastes the score column 
/// from each evaluation sheet into the summary sheet (for each team). 
/// 
/// Add the Reference to Microsoft.Office.Interop.Excel.
/// 

using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassAssessments
{
  class ParseWorkbooks
  {
  
    static void Main( string[] args )
    {
      if ( args.Length < 1 || args.Length > 1 )
      {
        Console.WriteLine( "Useage: " );
        Console.WriteLine( @"C:\UWT\ClassAssessments Team" );
        Console.WriteLine( "where Team is the team number to assess" );
      }
      else
      {
        ParseWorkbooks pwb = new ParseWorkbooks( args[0] );
      }
      Console.WriteLine( "\nAll done!\nHit Return to exit." );
      Console.ReadLine();
    }  // end of Main()


    /// <summary>
    /// ParseWrokbooks constructor does all the work.
    /// </summary>
    /// <param name="teamNumberString"></param>
    /// 
    ParseWorkbooks( string teamNumberString )
    {
      int teamNumber = TeamWorkbook.GetTeamNumber( teamNumberString );

      DirectoryInfo topDir = new DirectoryInfo( Environment.CurrentDirectory );  // directory we are in
      string topDirName = topDir.FullName;                                       // full path 
      string teamWorkbookPath = TeamWorkbook.GetTeamPath( topDirName, teamNumberString );  // make the path to our team workbook

      Excel.Application excelApp = new Excel.Application();  // Creates a new Excel Application
      excelApp.Visible = false;                              // Makes Excel invisible to the user.

      TeamWorkbook twb = new TeamWorkbook( excelApp, teamWorkbookPath ); // open our team workbook (must use full path)

      FileInfo[] studentFiles = topDir.GetFiles( "*.xlsx" );   // find all the student assessments
      if ( studentFiles.Length != 0 )
      {
        foreach ( FileInfo fileinf in studentFiles )  // for each student assessment
        {
          StudentWorkbook swb = new StudentWorkbook( excelApp, fileinf.FullName );  // parse the Excel file
          if ( swb.TeamNumber == teamNumber )   // if it's the correct team number
          {
            Console.WriteLine( fileinf.Name );  
            twb.PasteScores( swb.Scores );      // paste scores in the team workbook
            twb.PasteComments( swb.Comments );  // paste comments in teh team workbook
          }
          swb.Close();  // close the student assessment
        }
      }

      twb.Close();                                 // close the team workbook
      excelApp.Application.DisplayAlerts = false;  // no nanny
      excelApp.Quit();                             // close the Excel app

    }  // end of ParseWorkbooks
  }  // end of class ParseWorkbooks
}  // end of namespace ClassAssessments

