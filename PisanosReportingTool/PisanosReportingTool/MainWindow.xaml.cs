using System;
using System.Globalization;
using System.Windows;
using ExcelApi;
using ExcelApi.Models;
using Window = System.Windows.Window;

namespace PisanosReportingTool.ui
{
  public partial class MainWindow : Window
  {
    public MainWindow()
    {
      InitializeComponent();
    }

    private void LoadDailySummaryClick(object sender, RoutedEventArgs e)
    {
      var newDailySummaryLoader = new NewDailySummaryLoader();
      newDailySummaryLoader.GetFileNameFromUser();

      var fileNameReceivedFromUser = newDailySummaryLoader.FileName.Length > 0;
      if (!fileNameReceivedFromUser) return;

      var dailySummary = newDailySummaryLoader.ImportDailySummary();
      SetUiValuesForLoadedDailySummary(dailySummary);
    }

    private void SetUiValuesForLoadedDailySummary(DailySummary dailySummary)
    {
      FoodBevLunchTb.Text = dailySummary.SalesComparison.NetFoodBeverageSalesLunch.ToString(CultureInfo.InvariantCulture);
      FoodBevDinnerTb.Text = dailySummary.SalesComparison.NetFoodBeverageSalesDinner.ToString(CultureInfo.InvariantCulture);
      AlcoholLunchTb.Text = dailySummary.SalesComparison.NetAlcoholSalesLunch.ToString(CultureInfo.InvariantCulture);
      AlcoholDinnerTb.Text = dailySummary.SalesComparison.NetAlcoholSalesDinner.ToString(CultureInfo.InvariantCulture);
      OnlineSalesTb.Text = dailySummary.SalesComparison.NetOnlineSales.ToString(CultureInfo.InvariantCulture);
      CateringSales.Text = dailySummary.SalesComparison.NetCateringSales.ToString(CultureInfo.InvariantCulture);
      LunchCoversTb.Text = dailySummary.Covers.LunchCovers.ToString(CultureInfo.InvariantCulture);
      DinnerCoversTb.Text = dailySummary.Covers.DinnerCovers.ToString(CultureInfo.InvariantCulture);
      CashDepositTb.Text = dailySummary.Cash.CashDeposits.ToString(CultureInfo.InvariantCulture);
      OverShortTb.Text = dailySummary.Cash.OverShort.ToString(CultureInfo.InvariantCulture);
      PaidOutTb.Text = dailySummary.Cash.PaidOuts.ToString(CultureInfo.InvariantCulture);
      GiftCardsRedeemedTb.Text = dailySummary.Cash.GiftCardsRedeemed.ToString(CultureInfo.InvariantCulture);
      EightySixTb.Text = dailySummary.FoodVoids.EightySix.ToString(CultureInfo.InvariantCulture);
      CanceledOrderTb.Text = dailySummary.FoodVoids.CanceledOrder.ToString(CultureInfo.InvariantCulture);
      TrainingTb.Text = dailySummary.FoodVoids.Training.ToString(CultureInfo.InvariantCulture);
      ChangedMindTb.Text = dailySummary.FoodVoids.ChangedMind.ToString(CultureInfo.InvariantCulture);
      ServerErrorTb.Text = dailySummary.FoodVoids.ServerError.ToString(CultureInfo.InvariantCulture);
      ManagerMealTb.Text = dailySummary.FoodComps.ManagerMeal.ToString(CultureInfo.InvariantCulture);
      OwnerTb.Text = dailySummary.FoodComps.Owner.ToString(CultureInfo.InvariantCulture);
      DrawerMealTb.Text = dailySummary.FoodComps.DrawerMeal.ToString(CultureInfo.InvariantCulture);
      DonationTb.Text = dailySummary.FoodComps.Donation.ToString(CultureInfo.InvariantCulture);
      EmployeeOnShiftTb.Text = dailySummary.FoodDiscounts.EmployeeOnShift.ToString(CultureInfo.InvariantCulture);
      EmployeeOffShiftTb.Text = dailySummary.FoodDiscounts.EmployeeOffShift.ToString(CultureInfo.InvariantCulture);
      BdayAnniversaryTb.Text = dailySummary.FoodDiscounts.BirthdayAnniversary.ToString(CultureInfo.InvariantCulture);
      PromotionAdTb.Text = dailySummary.FoodDiscounts.PromotionAd.ToString(CultureInfo.InvariantCulture);
      MilitaryTb.Text = dailySummary.FoodDiscounts.Military.ToString(CultureInfo.InvariantCulture);
      FirePoliceTb.Text = dailySummary.FoodDiscounts.FirePolice.ToString(CultureInfo.InvariantCulture);
      GoodCustomerTb.Text = dailySummary.FoodDiscounts.GoodCustomer.ToString(CultureInfo.InvariantCulture);
      CityOfKennesawTb.Text = dailySummary.FoodDiscounts.CityOfKennesaw.ToString(CultureInfo.InvariantCulture);
      CobbTeacherTb.Text = dailySummary.FoodDiscounts.CobbTeachers.ToString(CultureInfo.InvariantCulture);
      OtherRestaurantTb.Text = dailySummary.FoodDiscounts.OtherRestaurant.ToString(CultureInfo.InvariantCulture);
      ManagerOwnerTb.Text = dailySummary.FoodDiscounts.OwnerManager.ToString(CultureInfo.InvariantCulture);
    }

    private void SaveDailySummaryButtonClick(object sender, RoutedEventArgs e)
    {

    }
  }
}