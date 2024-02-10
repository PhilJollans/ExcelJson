using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.Security.Cryptography;

namespace ExcelJson
{
  public partial class ExcelJsonRibbon
  {

    private void ReadButton_Click( object sender, RibbonControlEventArgs e )
    {
      try
      {
        var ofd = new OpenFileDialog
        {
          Title = "Select json file",
          DefaultExt = "json",
          Filter = "json files (*.json)|*.json|All Files (*.*)|*.*",
          CheckFileExists = true,
        } ;

        var response = ofd.ShowDialog();
        if ( response == DialogResult.OK )
        {
          // Read the file
          string jsonString = File.ReadAllText(ofd.FileName);

          // Parse into a JToken
          JToken rootToken = JToken.Parse(jsonString);

          // Check if the root object is an array
          if ( rootToken.Type != JTokenType.Array )
          {
            MessageBox.Show ( "root element of json file must be an array.", "ExcelJson", MessageBoxButtons.OK, MessageBoxIcon.Error );
            return;
          }

          // Cast to JArray
          JArray jsonArrayObject = rootToken as JArray ;

          // Get the first element in the array to extract property names
          JObject firstObject = (JObject)jsonArrayObject.First;

          // Build a C# array with property names
          string[] propertyNames = firstObject.Properties().Select(p => p.Name).ToArray();

          // Get the base name of the file and create a new worksheet.
          var basename = Path.GetFileNameWithoutExtension ( ofd.FileName ) ;
          var worksheet = GetWorksheetForBaseName ( basename ) ;

          // Write the header row
          int row = 1 ;
          for ( int col = 0 ; col < propertyNames.Length ; col++ )
          {
            worksheet.Cells[row, col+1] = propertyNames[col];
          }

          // Iterate through the objects
          foreach ( JObject jsonObject in jsonArrayObject.Children<JObject> () )
          {
            row++ ;

            // Fetch properties based on the extracted property names
            for ( int col = 0 ; col < propertyNames.Length ; col++ )
            {
              var propertyName = propertyNames[col];
              // Get the property value for the current property name
              JToken propertyValue = jsonObject[propertyName];
              worksheet.Cells[row, col+1] = propertyValue ;
            }
          }

          // Adjust the column sizes
          worksheet.Columns.AutoFit() ;

          // Set the background colour for the header cells.
          Range headerRowRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, propertyNames.Length]];
          headerRowRange.Interior.Color = XlRgbColor.rgbLightBlue;
        }
      }
      catch ( Exception ex )
      {
        MessageBox.Show ( $"Exception {ex.GetType ().ToString ()}, {ex.Message} in ExcelJson" ) ;
      }
    }

    private void WriteButton_Click( object sender, RibbonControlEventArgs e )
    {
      try
      {
        var sfd = new SaveFileDialog
        {
          Title = "Select json file",
          DefaultExt = "json",
          Filter = "json files (*.json)|*.json|All Files (*.*)|*.*"
        } ;

        var response = sfd.ShowDialog();
        if ( response == DialogResult.OK )
        {
          // Get the active worksheet
          Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet ;

          // Get the used range
          var usedRange = worksheet.UsedRange ;


          // Determine how many columns there are.

        }
      }
      catch ( Exception ex )
      {
        MessageBox.Show ( $"Exception {ex.GetType ().ToString ()}, {ex.Message} in ExcelJson" );
      }
    }

    private Worksheet GetWorksheetForBaseName( string BaseName )
    {
      // Look for a worksheet with the basename of file
      Workbook Book = Globals.ThisAddIn.Application.ActiveWorkbook;

      // Look for an existing sheet with this name
      foreach ( Worksheet Sheet in Book.Sheets )
      {
        if ( Sheet.Name.Equals ( BaseName, StringComparison.OrdinalIgnoreCase ) )
        {
          return Sheet;
        }
      }

      // If we didn't find it, then create a new worksheet
      var ResSheet = Book.Worksheets.Add();
      ResSheet.Name = BaseName;

      return ResSheet;
    }
  }
}
