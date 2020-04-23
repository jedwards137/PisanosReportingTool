using Microsoft.Win32;
using System;
using System.Windows;
using Window = System.Windows.Window;

namespace PisanosReportingTool
{
  public partial class MainWindow : Window
  {
    public MainWindow()
    {
      InitializeComponent();
    }

    private void Button_Click(object sender, RoutedEventArgs e)
    {
      var dailySummaryFilename = GetDailySummaryFilenameFromUser();
      var hasFileName = dailySummaryFilename != string.Empty;
      if (!hasFileName) return;

      var loadedDailySummarySheet = new DailySummarySheet(dailySummaryFilename);
      setUiValuesForLoadedDailySummarySheet(loadedDailySummarySheet);
      Console.WriteLine("");
    }

    private void setUiValuesForLoadedDailySummarySheet(DailySummarySheet loadedDailySummarySheet)
    {
      
    }

    private static string GetDailySummaryFilenameFromUser()
    {
      var openFileDialog = new OpenFileDialog { Filter = "Excel Office | *.xlsx;*.xls" };
      openFileDialog.ShowDialog();
      var fileName = openFileDialog.FileName;
      return fileName;
    }
  }
}