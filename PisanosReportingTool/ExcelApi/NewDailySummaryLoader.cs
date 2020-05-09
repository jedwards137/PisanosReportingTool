using System;
using ExcelApi.Models;
using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelApi
{
  public class NewDailySummaryLoader
  {
    public string FileName { get; private set; }

    public void GetFileNameFromUser()
    {
      var openFileDialog = new OpenFileDialog { Filter = "Excel Office | *.xlsx;*.xls" };
      openFileDialog.ShowDialog();
      var fileName = openFileDialog.FileName;
      FileName = fileName;
    }

    public DailySummary ImportDailySummary()
    {
      var fileNameNotFound = !(FileName.Length > 0);
      if (fileNameNotFound) return null;

      var rawSummaryData = GetRawSummaryDataFromFile(FileName);
      var categoryRowIndexes = GetCategoryRowIndexes(rawSummaryData);

      var dailySummary = new DailySummary
      {
        Date = GetDate(rawSummaryData),
        SalesComparison = GetSalesComparisonValues(rawSummaryData, categoryRowIndexes),
        Covers = GetCoversValues(rawSummaryData, categoryRowIndexes),
        Cash = GetCashValues(rawSummaryData, categoryRowIndexes),
        FoodVoids = GetFoodVoids(rawSummaryData, categoryRowIndexes),
        FoodComps = GetFoodComps(rawSummaryData, categoryRowIndexes),
        FoodDiscounts = GetFoodDiscounts(rawSummaryData, categoryRowIndexes)
      };


      return dailySummary;
    }

    private static FoodDiscounts GetFoodDiscounts(DataTable rawSummaryData, IReadOnlyDictionary<string, int> categoryRowIndexes)
    {
      var foodDiscounts = new FoodDiscounts();

      var startingIndex = categoryRowIndexes["DiscountSummary"];
      var endingIndex = categoryRowIndexes["GuestTableInformationByPeriod"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = rawSummaryData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString().Trim().ToLower();

        switch (rowTitle)
        {
          case "cobb county teacher":
            var cobbTeachers = (double)currentRow[5];
            foodDiscounts.CobbTeachers = cobbTeachers;
            break;
          case "employee":
            var employeeOnShift = (double)currentRow[5];
            foodDiscounts.EmployeeOnShift = employeeOnShift;
            break;
          case "employee off shift":
            var employeeOffShift = (double)currentRow[5];
            foodDiscounts.EmployeeOffShift = employeeOffShift;
            break;
          case "fire dept":
            var fire = (double)currentRow[5];
            foodDiscounts.FirePolice += fire;
            break;
          case "police":
            var police = (double)currentRow[5];
            foodDiscounts.FirePolice += police;
            break;
          case "good customer":
            var goodCustomer = (double)currentRow[5];
            foodDiscounts.GoodCustomer = goodCustomer;
            break;
          case "in-house promo":
            var promotionAd = (double)currentRow[5];
            foodDiscounts.PromotionAd = promotionAd;
            break;
          case "manager":
            var ownerManager = (double)currentRow[5];
            foodDiscounts.OwnerManager = ownerManager;
            break;
        }
      }

      return foodDiscounts;
    }

    private static FoodComps GetFoodComps(DataTable rawSummaryData, IReadOnlyDictionary<string, int> categoryRowIndexes)
    {
      var foodComps = new FoodComps();

      var startingIndex = categoryRowIndexes["CompSummary"];
      var endingIndex = categoryRowIndexes["DiscountSummary"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = rawSummaryData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString().Trim().ToLower();

        switch (rowTitle)
        {
          case "manager meal":
            var managerMeal = (double)currentRow[5];
            foodComps.ManagerMeal = managerMeal;
            break;
          case "drawer meal":
            var drawerMeal = (double)currentRow[5];
            foodComps.DrawerMeal = drawerMeal;
            break;
          case "donation":
            var donation = (double)currentRow[5];
            foodComps.Donation = donation;
            break;
        }
      }

      return foodComps;
    }

    private static FoodVoids GetFoodVoids(DataTable rawSummaryData, IReadOnlyDictionary<string, int> categoryRowIndexes)
    {
      var foodVoids = new FoodVoids();

      var startingIndex = categoryRowIndexes["VoidSummary"];
      var endingIndex = categoryRowIndexes["CompSummary"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = rawSummaryData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString().Trim().ToLower();

        switch (rowTitle)
        {
          case "changed mind":
            var changedMind = (double)currentRow[5];
            foodVoids.ChangedMind = changedMind;
            break;
          case "86":
            var eightysix = (double)currentRow[5];
            foodVoids.EightySix = eightysix;
            break;
          case "canceled order":
            var canceledOrder = (double)currentRow[5];
            foodVoids.CanceledOrder = canceledOrder;
            break;
          case "server error":
            var serverError = (double)currentRow[5];
            foodVoids.ServerError = serverError;
            break;
          case "training":
            var training = (double)currentRow[5];
            foodVoids.Training = training;
            break;
        }
      }

      return foodVoids;
    }

    private static Cash GetCashValues(DataTable rawSummaryData, IReadOnlyDictionary<string, int> categoryRowIndexes)
    {
      var cash = new Cash();

      var startingIndex = categoryRowIndexes["Deposit"];
      var endingIndex = rawSummaryData.Rows.Count;

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = rawSummaryData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString().Trim().ToLower();

        switch (rowTitle)
        {
          case "net cash received":
            var cashDeposit = Convert.ToDouble(currentRow[4]);
            cash.CashDeposits = cashDeposit;
            break;
          case "paid out":
            var paidOut = Convert.ToDouble(currentRow[4]);
            cash.PaidOuts = paidOut;
            break;
        }
      }

      return cash;
    }

    private static Covers GetCoversValues(DataTable rawSummaryData, IReadOnlyDictionary<string, int> categoryRowIndexes)
    {
      var covers = new Covers();

      var startingIndex = categoryRowIndexes["GuestTableInformationByPeriod"];
      var endingIndex = categoryRowIndexes["GuestTableInformationByProfitCenter"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = rawSummaryData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString().Trim().ToLower();

        var coversRowFound = rowTitle == "number guest served";
        if (!coversRowFound) continue;

        var lunchCovers = (double)currentRow[1];
        covers.LunchCovers = lunchCovers;

        var dinnerCovers = (double)currentRow[2];
        covers.DinnerCovers = dinnerCovers;
      }

      return covers;
    }

    private static SalesComparison GetSalesComparisonValues(DataTable rawSummaryData, IReadOnlyDictionary<string, int> categoryRowIndexes)
    {
      var salesComparison = new SalesComparison();

      var startingIndex = categoryRowIndexes["SummaryByPeriod"];
      var endingIndex = categoryRowIndexes["SummaryByProfitCenter"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = rawSummaryData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString().Trim().ToLower();

        string[] lunchDinnerTitles = {"beverage", "delivery", "food"};
        var rowTitleMatchesLunchDinnerTitle = lunchDinnerTitles.Contains(rowTitle);
        if (rowTitleMatchesLunchDinnerTitle)
        {
          var foodBevLunchValue = (double) currentRow[1];
          salesComparison.NetFoodBeverageSalesLunch += foodBevLunchValue;

          var foodBevDinnerValue = (double) currentRow[2];
          salesComparison.NetFoodBeverageSalesDinner += foodBevDinnerValue;
        }

        var rowTitleMatchesAlcohol = rowTitle == "alcohol";
        if (!rowTitleMatchesAlcohol) continue;
        var alcoholLunchValue = (double) currentRow[1];
        salesComparison.NetAlcoholSalesLunch = alcoholLunchValue;

        var alcoholDinnerValue = (double) currentRow[2];
        salesComparison.NetAlcoholSalesDinner = alcoholDinnerValue;
      }

      startingIndex = categoryRowIndexes["PaymentSummary"];
      endingIndex = categoryRowIndexes["SummaryByPeriod"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = rawSummaryData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString().Trim().ToLower();

        string[] onlineTitles = {"grubhub", "doordash", "eatstreet", "slice", "menufy"};
        var rowTitleMatchesOnlineTitle = onlineTitles.Contains(rowTitle);
        if (!rowTitleMatchesOnlineTitle) continue;
        var onlineValue = (double) currentRow[5];
        salesComparison.NetOnlineSales += onlineValue;
      }

      return salesComparison;
    }

    private static DateTime GetDate(DataTable rawSummaryData)
    {
      var firstCell = rawSummaryData.Rows[0].ItemArray[0].ToString()?.ToLower() ?? "";
      var dateString = firstCell.Split(' ').Last();
      var dateValues = dateString.Split('/');

      var year = 2000 + Convert.ToInt32(dateValues[2]);
      var month = Convert.ToInt32(dateValues[0]);
      var day = Convert.ToInt32(dateValues[1]);
      var date = new DateTime(year, month, day);

      return date;
    }

    private static DataTable GetRawSummaryDataFromFile(string fileName)
    {
      Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
      var stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
      var reader = ExcelReaderFactory.CreateReader(stream);
      var dataSet = reader.AsDataSet();
      var rawSheetData = dataSet.Tables[0];

      return rawSheetData;
    }

    private static Dictionary<string, int> GetCategoryRowIndexes(DataTable rawSummaryData)
    {
      var categoryRowIndexes = new Dictionary<string, int>();

      for (var i = 0; i < rawSummaryData.Rows.Count; i++)
      {
        foreach (var item in rawSummaryData.Rows[i].ItemArray)
        {
          var data = item + "";
          data = data.ToLower();

          switch (data)
          {
            case "category summary by period":
              categoryRowIndexes.Add("SummaryByPeriod", i);
              break;
            case "category summary by profit center":
              categoryRowIndexes.Add("SummaryByProfitCenter", i);
              break;
            case "payment summary":
              categoryRowIndexes.Add("PaymentSummary", i);
              break;
            case "guest and table information by period":
              categoryRowIndexes.Add("GuestTableInformationByPeriod", i);
              break;
            case "deposit":
              categoryRowIndexes.Add("Deposit", i);
              break;
            case "void summary":
              categoryRowIndexes.Add("VoidSummary", i);
              break;
            case "comp summary":
              categoryRowIndexes.Add("CompSummary", i);
              break;
            case "discount summary":
              categoryRowIndexes.Add("DiscountSummary", i);
              break;
            case "guest and table information by profit center":
              categoryRowIndexes.Add("GuestTableInformationByProfitCenter", i);
              break;
            case null:
              break;
          }
        }
      }

      return categoryRowIndexes;
    }
  }
}
