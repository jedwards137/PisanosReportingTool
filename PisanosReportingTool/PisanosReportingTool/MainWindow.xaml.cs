using Microsoft.Win32;
using System;
using System.Globalization;
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

    private void Button_Click_1(object sender, RoutedEventArgs e)
    {
      var dailySummaryFilename = GetDailySummaryFilenameFromUser();
      var hasFileName = dailySummaryFilename != string.Empty;
      if (!hasFileName) return;

      var loadedDailySummarySheet = new DailySummarySheet(dailySummaryFilename);
      SetUiValuesForLoadedDailySummarySheet(loadedDailySummarySheet);
      Console.WriteLine("");
    }

    private void SetUiValuesForLoadedDailySummarySheet(DailySummarySheet loadedDailySummarySheet)
    {
      FoodBevLunchTb.Text = loadedDailySummarySheet.ParsedSheetData["FoodBevLunch"].ToString(CultureInfo.InvariantCulture);
      FoodBevDinnerTb.Text = loadedDailySummarySheet.ParsedSheetData["FoodBevDinner"].ToString(CultureInfo.InvariantCulture);
      AlcoholLunchTb.Text = loadedDailySummarySheet.ParsedSheetData["AlcoholLunch"].ToString(CultureInfo.InvariantCulture);
      AlcoholDinnerTb.Text = loadedDailySummarySheet.ParsedSheetData["AlcoholDinner"].ToString(CultureInfo.InvariantCulture);
      OnlineSalesTb.Text = loadedDailySummarySheet.ParsedSheetData["OnlineSales"].ToString(CultureInfo.InvariantCulture);
      CateringSales.Text = loadedDailySummarySheet.ParsedSheetData["CateringSales"].ToString(CultureInfo.InvariantCulture);
      LunchCoversTb.Text = loadedDailySummarySheet.ParsedSheetData["LunchCovers"].ToString(CultureInfo.InvariantCulture);
      DinnerCoversTb.Text = loadedDailySummarySheet.ParsedSheetData["DinnerCovers"].ToString(CultureInfo.InvariantCulture);
      CashDepositTb.Text = loadedDailySummarySheet.ParsedSheetData["CashDeposit"].ToString(CultureInfo.InvariantCulture);
      OverShortTb.Text = loadedDailySummarySheet.ParsedSheetData["OverShort"].ToString(CultureInfo.InvariantCulture);
      PaidOutTb.Text = loadedDailySummarySheet.ParsedSheetData["PaidOut"].ToString(CultureInfo.InvariantCulture);
      GiftCardsRedeemedTb.Text = loadedDailySummarySheet.ParsedSheetData["GiftCardsRedeemed"].ToString(CultureInfo.InvariantCulture);
      EightySixTb.Text = loadedDailySummarySheet.ParsedSheetData["86"].ToString(CultureInfo.InvariantCulture);
      CanceledOrderTb.Text = loadedDailySummarySheet.ParsedSheetData["CanceledOrder"].ToString(CultureInfo.InvariantCulture);
      TrainingTb.Text = loadedDailySummarySheet.ParsedSheetData["Training"].ToString(CultureInfo.InvariantCulture);
      ChangedMindTb.Text = loadedDailySummarySheet.ParsedSheetData["ChangedMind"].ToString(CultureInfo.InvariantCulture);
      ServerErrorTb.Text = loadedDailySummarySheet.ParsedSheetData["ServerError"].ToString(CultureInfo.InvariantCulture);
      ManagerMealTb.Text = loadedDailySummarySheet.ParsedSheetData["ManagerMeal"].ToString(CultureInfo.InvariantCulture);
      OwnerTb.Text = loadedDailySummarySheet.ParsedSheetData["Owner"].ToString(CultureInfo.InvariantCulture);
      DrawerMealTb.Text = loadedDailySummarySheet.ParsedSheetData["DrawerMeal"].ToString(CultureInfo.InvariantCulture);
      DonationTb.Text = loadedDailySummarySheet.ParsedSheetData["Donation"].ToString(CultureInfo.InvariantCulture);
      EmployeeOnShiftTb.Text = loadedDailySummarySheet.ParsedSheetData["EmployeeOnShift"].ToString(CultureInfo.InvariantCulture);
      EmployeeOffShiftTb.Text = loadedDailySummarySheet.ParsedSheetData["EmployeeOffShift"].ToString(CultureInfo.InvariantCulture);
      BdayAnniversaryTb.Text = loadedDailySummarySheet.ParsedSheetData["BdayAnniversary"].ToString(CultureInfo.InvariantCulture);
      PromotionAdTb.Text = loadedDailySummarySheet.ParsedSheetData["InHousePromo"].ToString(CultureInfo.InvariantCulture);
      MilitaryTb.Text = loadedDailySummarySheet.ParsedSheetData["Military"].ToString(CultureInfo.InvariantCulture);
      FirePoliceTb.Text = loadedDailySummarySheet.ParsedSheetData["FirePolice"].ToString(CultureInfo.InvariantCulture);
      GoodCustomerTb.Text = loadedDailySummarySheet.ParsedSheetData["GoodCustomer"].ToString(CultureInfo.InvariantCulture);
      CityOfKennesawTb.Text = loadedDailySummarySheet.ParsedSheetData["CityOfKennesaw"].ToString(CultureInfo.InvariantCulture);
      CobbTeacherTb.Text = loadedDailySummarySheet.ParsedSheetData["CobbTeacher"].ToString(CultureInfo.InvariantCulture);
      OtherRestaurantTb.Text = loadedDailySummarySheet.ParsedSheetData["OtherRestaurant"].ToString(CultureInfo.InvariantCulture);
      ManagerOwnerTb.Text = loadedDailySummarySheet.ParsedSheetData["ManagerOwner"].ToString(CultureInfo.InvariantCulture);
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