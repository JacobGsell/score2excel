using System;
using static System.DateTime;
using System.Collections.Generic;
using System.Windows;
using static System.IO.Directory;
using static System.Runtime.InteropServices.Marshal;
using Microsoft.Office.Interop.Excel;

namespace wpfChemie
{
  /// <summary>
  /// Interaktionslogik für MainWindow.xaml
  /// </summary>
  public partial class MainWindow : System.Windows.Window
  {
    int selectTable;  // Decides which table to choose from

    /// <summary>
    /// Convert an excel range into an array
    /// </summary>
    /// <param name="r">Selected range</param>
    /// <returns>Array</returns>
    static string[,] range2Array(Range r)
    {
      string[,] arr = new string[r.Cells.Rows.Count, r.Cells.Columns.Count];

      for (int i = 1; i <= r.Rows.Count; i++)
      {
        for (int j = 1; j <= r.Columns.Count; j++)
        {
          arr[i - 1, j - 1] = r.Cells[i, j].Value.ToString();
        }
      }
      return arr;
    }

    /// <summary>
    /// Get x axis for value
    /// </summary>
    /// <param name="matrix">Excel table to search in</param>
    /// <param name="c">Current character in string</param>
    /// <returns>X axis</returns>
    static int getCharPos(string[,] matrix, char c)
    {
      for (int i = 0; i < matrix.GetLength(0); i++)
      {
        if (matrix[i, 0] == c.ToString())
        {
          return i;
        }
      }
      return 0;
    }

    /// <summary>
    /// Get the y axis of the demanded value
    /// </summary>
    /// <param name="matrix">Excel table to search in</param>
    /// <param name="pos">Current 'P' of Value</param>
    /// <returns>Y axis</returns>
    static int getPositionPos(string[,] matrix, string pos)
    {
      for (int i = 0; i < matrix.GetLength(1); i++)
      {
        if (matrix[0, i] == pos)
        {
          return i;
        }
      }
      return 0;
    }

    /// <summary>
    /// Get the position of value in table
    /// </summary>
    /// <param name="matrix">Excel table with values</param>
    /// <param name="pos">Value between P4 and P4' in</param>
    /// <param name="c">Specific character in string</param>
    /// <returns>Value</returns>
    static int getPosition(string[,] matrix, string pos, char c)
    {
      int x = getCharPos(matrix, c);
      int y = getPositionPos(matrix, pos);

      return Convert.ToInt32(matrix[x, y].ToString());
    }

    /// <summary>
    /// Add a chart to the table
    /// </summary>
    /// <param name="ws">Current worksheet</param>
    /// <param name="scores">List of calculated scores</param>
    /// <param name="chartType">Type of chart</param>
    static void addChart(Worksheet ws, List<int> scores, XlChartType chartType)
    {
      var charts = ws.ChartObjects() as ChartObjects;
      var chartObject = charts.Add(0, 50, 800, 300) as ChartObject;
      var chart = chartObject.Chart;
      var range = ws.Range[ws.Cells[1, 1], ws.Cells[3, scores.Count]];

      // Set up the chart:
      chart.SetSourceData(range);
      chart.ChartType = chartType;
      chart.HasLegend = false;
      chart.HasTitle = true;
      chart.ChartTitle.Text = "All Scores";

      // Define the content of the x axis:
      Range xValRange = ws.Range[ws.Cells[1, 1], ws.Cells[2, scores.Count]];
      Axis xAxis = (Axis)chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
      xAxis.CategoryNames = xValRange;
    }

    /// <summary>
    /// Create excel file
    /// </summary>
    /// <param name="scores">List of calculated scores</param>
    /// <param name="input">Input string</param>
    /// <param name="old">Time from when the Button was clicked</param>
    /// <param name="setChart">Display or hide chart</param>
    static void exportToExcel(List<int> scores, string input, DateTime old, bool setChart = false)
    {
      var myApp = new Microsoft.Office.Interop.Excel.Application();

      if (myApp == null)
      {
        return;
      }

      Workbook wb;
      Worksheet ws;

      object misValue = System.Reflection.Missing.Value;

      wb = myApp.Workbooks.Add(misValue);
      ws = wb.Worksheets.get_Item(1);

      for (int i = 1; i <= scores.Count; i++)
      {
        ws.Cells[1, i] = "[" + (i + 3).ToString() + "]";
        ws.Cells[2, i] = input[i + 2].ToString();
        ws.Cells[3, i] = scores[i - 1].ToString();
      }

      if (setChart)
      {
        addChart(ws, scores, XlChartType.xlColumnStacked);
      }

      DateTime now = Now;
      string fileNameTime = now.Hour.ToString() + "-" + now.Minute.ToString() + "-" + now.Second.ToString();

      // Add scores folder if it doesnt exist yet:
      var path = GetCurrentDirectory();
      if (!Exists(path + @"\Scores"))
      {
        CreateDirectory((path + @"\Scores"));
      }

      // Save excel file into Scores folder:
      path = path + @"\Scores" + @"\" + Now.ToShortDateString() + "; " + fileNameTime + ".xlsx";
      wb.SaveAs(path, XlFileFormat.xlOpenXMLWorkbook,
      misValue, misValue, misValue, misValue,
      XlSaveAsAccessMode.xlExclusive,
      misValue, misValue, misValue, misValue, misValue);
      wb.Close(true, misValue, misValue);
      myApp.Quit();

      // Release the file:
      ReleaseComObject(ws);
      ReleaseComObject(wb);
      ReleaseComObject(myApp);

      // Show time the program took to finish:
      MessageBox.Show("Excel Datei wurde in " +
          string.Format("{0:0.##}", (Now - old).TotalSeconds) +
         " Sekunden erfolgreich erstellt.");

      // Open the folder which contains the files:
      System.Diagnostics.Process.Start(GetCurrentDirectory() + @"\Scores");
    }

    public MainWindow()
    {
      InitializeComponent();
    }

    private void createExcel(object sender, RoutedEventArgs e)
    {
      DateTime old = Now;  //  Measuring time

      string[] positions = { "P4", "P3", "P2", "P1", "P1'", "P2'", "P3'", "P4'" };

      // Get table:
      var app = new Microsoft.Office.Interop.Excel.Application();
      var path = GetCurrentDirectory();

      // Open table:
      Workbook workbook = app.Workbooks.Open(@"" + path + @"\matrizen_test.xlsx");
      Worksheet worksheet = workbook.ActiveSheet;

      // Create array from Matrix and close file afterwards:
      string[,] matrix = getMatrix(worksheet);
      workbook.Close(true, null, null);
      app.Quit();

      string input = this.input.Text.ToUpper();

      List<int> scores = new List<int>();

      // Go through the string:
      int sum;
      for (int i = 3; i < input.Length - 4; i++)
      {
        sum = 0;

        // Add the values to get a score:
        for (int j = 0; j < 8; j++)
        {
          sum += getPosition(matrix, positions[j], input[i - 3 + j]);
        }
        scores.Add(sum);
      }

      exportToExcel(scores, input, old, true);
    }

    private void setCATD(object sender, RoutedEventArgs e)
    {
      selectTable = 0;
    }

    private void setPlasmin(object sender, RoutedEventArgs e)
    {
      selectTable = 1;
    }

    private void setThrombin(object sender, RoutedEventArgs e)
    {
      selectTable = 2;
    }

    // Check if string is valid:
    private void checkString(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
      var r = new System.Text.RegularExpressions.Regex("[^G,P,A,V,L,I,M,F,Y,W,S,T,C,N,Q,D,E,K,R,H]");
      if (r.IsMatch(input.Text.ToUpper()) || input.Text == string.Empty)
      {
        exportButton.IsEnabled = false;
      }

      else
      {
        exportButton.IsEnabled = true;
      }
    }

    private string[,] getMatrix(Worksheet ws)
    {
      string[,] matrix = null;

      switch (selectTable)
      {
        case 0:
          matrix = range2Array(ws.Range["B4", "J24"]);  // CATD
          break;
        case 1:
          matrix = range2Array(ws.Range["B27", "J47"]);  // Plasmin
          break;
        case 2:
          matrix = range2Array(ws.Range["B51", "J71"]);  // Thrombin
          break;
      }
      return matrix;
    }
  }
}
