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
    //string topDirName;  // name of the start-up directory

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


    ParseWorkbooks( string teamNumberString )
    {
      int teamNumber = TeamWorkbook.GetTeamNumber( teamNumberString );

      DirectoryInfo topDir = new DirectoryInfo( Environment.CurrentDirectory );
      string topDirName = topDir.FullName;
      string teamWorkbookPath = TeamWorkbook.GetTeamPath( topDirName, teamNumberString );

      Excel.Application excelApp = new Excel.Application();  // Creates a new Excel Application
      excelApp.Visible = false;  // Makes Excel invisible to the user.

      //TeamWorkbook twb = new TeamWorkbook( excelApp, "e:\\projects\\cs490\\excel\\classassessments\\bin\\debug\\Team01\\Team01P1.xlsx" );
      TeamWorkbook twb = new TeamWorkbook( excelApp, teamWorkbookPath );

      FileInfo[] studentFiles = topDir.GetFiles( "*.xlsx" );
      if ( studentFiles.Length != 0 )
      {
        foreach ( FileInfo fileinf in topDir.GetFiles( "*.xlsx" ) )
        //FileInfo fileinf = new FileInfo( "graydaniel_3484091_34593574_assessment1.xlsx" );
        {
          StudentWorkbook swb = new StudentWorkbook( excelApp, fileinf.FullName );
          if ( swb.TeamNumber == teamNumber )
          {
            Console.WriteLine( fileinf.Name );
            twb.PasteScores( swb.Scores );
            twb.PasteComments( swb.Comments );
          }
          swb.Close();
        }
      }

      twb.Close();
      excelApp.Application.DisplayAlerts = false;
      excelApp.Quit();

    }  // end of ParseWorkbooks
  }  // end of class ParseWorkbooks
}  // end of namespace ClassAssessments

