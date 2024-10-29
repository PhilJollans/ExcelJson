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

          // Build a list of the property names.
          List<string> propertyNames = firstObject.Properties().Select(p => p.Name).ToList();

          // Loop over the other objects and look for properties which are not in the first object.
          // Insert them into the list at the correct position.

          // Loop through remaining rows and insert unknown properties at the correct place.
          // This is not a watertight algorithm.
          foreach ( JObject jsonObject in jsonArrayObject.Skip (1).OfType<JObject> () )
          {
            var propertiesList = jsonObject.Properties().ToList();

            for ( int i = 0 ; i < propertiesList.Count ; i++ )
            {
              string propertyName = propertiesList[i].Name;

              if ( !propertyNames.Contains ( propertyName ) )
              {
                // If the property is not in the propertyOrderList, insert it after the preceding property
                if ( i == 0 )
                {
                  // Insert at the start of the list
                  propertyNames.Insert ( 0, propertyName ) ;
                }
                else
                {
                  string previousName = propertiesList[i-1].Name;
                  int precedingIndex = propertyNames.IndexOf(previousName);
                  propertyNames.Insert ( precedingIndex + 1, propertyName );
                }
              }
            }
          }

          // Get the base name of the file and create a new worksheet.
          var basename = Path.GetFileNameWithoutExtension ( ofd.FileName ) ;
          var worksheet = GetWorksheetForBaseName ( basename ) ;

          // Write the header row
          int row = 1 ;
          for ( int col = 0 ; col < propertyNames.Count ; col++ )
          {
            worksheet.Cells[row, col+1] = propertyNames[col];
          }

          // Iterate through the objects
          foreach ( JObject jsonObject in jsonArrayObject.Children<JObject> () )
          {
            row++ ;

            // Fetch properties based on the extracted property names
            for ( int col = 0 ; col < propertyNames.Count ; col++ )
            {
              var propertyName = propertyNames[col];

              // Get the property value for the current property name
              JToken propertyValue = jsonObject[propertyName];

              // Excel likes to detect a boolean value and convert it to a localized string.
              if ( ( propertyValue?.ToString() == "true" ) || ( propertyValue?.ToString() == "false" ) )
              {
                worksheet.Cells[row, col+1] = "'" + propertyValue ;
              }
              else
              {
                worksheet.Cells[row, col+1] = propertyValue ;
              }
            }
          }

          // Adjust the column sizes
          worksheet.Columns.AutoFit() ;

          // Set the background colour for the header cells.
          Range headerRowRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, propertyNames.Count]];
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

          int nRows = usedRange.Rows.Count;
          int nCols = usedRange.Columns.Count;

          // NOTE: I will use zero based indeces, even though excel uses 1 based indeces.
          // This means that we must add 1 when we access rows and cells.

          // Loop over the first row and get the property names
          //int nCols = usedRange.Rows[1].Cells.Count;
          string[] propertyNames = new string[nCols];
          for ( int col = 0 ; col < nCols ; col++ )
          {
            string cellValue = usedRange.Cells[1, col+1].Text;
            propertyNames[col] = cellValue;
          }

          // Initialize a list to store JSON objects
          List<JObject> jsonObjectsList = new List<JObject>();

          // Iterate through rows (excluding the header) and build JSON objects
          for ( int row = 1 ; row < nRows ; row++ ) // Start from 2 to skip the header row
          {
            JObject jsonObject = new JObject();

            // Iterate through each cell in the row and add properties to the JSON object
            for ( int col = 0 ; col < nCols ; col++ )
            {
              // Get the corresponding property name from the header row
              string propertyName = propertyNames[col];

              // Get the value from the current cell
              var cellValue = usedRange.Cells[row+1, col+1].Value;

              // Skip - don't export - null cells.
              if (  cellValue != null )
              {
                if ( cellValue is string cellString )
                {
                  // Prefer true and false in lower case.
                  if ( cellString.Equals ( "True", StringComparison.OrdinalIgnoreCase ) )
                  {
                    cellString = "true" ;
                  }
                  if ( cellString.Equals ( "False", StringComparison.OrdinalIgnoreCase ) )
                  {
                    cellString = "false" ;
                  }

                  // Add the property to the JSON object
                  jsonObject[propertyName] = JToken.FromObject ( cellString );
                }
#if false
                else if ( cellValue is bool cellBool )
                {
                  jsonObject[propertyName] = JToken.FromObject ( cellBool ? "true" : "false" );
                }
#endif
                else
                {
#if true
                  // Add the property to the JSON object
                  jsonObject[propertyName] = JToken.FromObject ( cellValue );
#else
                  jsonObject[propertyName] = JToken.FromObject ( cellValue.ToString() );
#endif
                }
              }

            }

            // Add the JSON object to the list
            jsonObjectsList.Add ( jsonObject );
          }

          // Serialize the list of JSON objects to a JSON string
          string jsonString = JsonConvert.SerializeObject(jsonObjectsList, Formatting.Indented);

          // And finally save it to a file
          File.WriteAllText ( sfd.FileName, jsonString );
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
