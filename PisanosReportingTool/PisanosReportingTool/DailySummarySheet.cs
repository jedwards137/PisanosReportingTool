using System;
using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace PisanosReportingTool
{
  public class DailySummarySheet
  {
    public Dictionary<string, double> ParsedSheetData { get; private set; }
    private readonly DataTable _rawSheetData;
    private readonly Dictionary<string, int> _categoryRowIndexes;

    public DailySummarySheet(string fileName)
    {
      ParsedSheetData = new Dictionary<string, double>();
      _rawSheetData = IngestRawDataFromExcelSheet(fileName);
      _categoryRowIndexes = FindCategoryRowIndexes();

      SetInitialValues();
      GatherSummaryByPeriodValues();
      GatherOnlineSalesValues();
      GatherCovers();
      GatherCashSection();
      GatherVoids();
      GatherComps();
      GatherDiscounts();
    }

    private void SetInitialValues()
    {
      ParsedSheetData["CateringSales"] = 0;
      ParsedSheetData["OverShort"] = 0;
      ParsedSheetData["GiftCardsRedeemed"] = 0;
      ParsedSheetData["Owner"] = 0;
      ParsedSheetData["BdayAnniversary"] = 0;
      ParsedSheetData["Military"] = 0;
      ParsedSheetData["CityOfKennesaw"] = 0;
      ParsedSheetData["OtherRestaurant"] = 0;

      ParsedSheetData["FoodBevLunch"] = 0;
      ParsedSheetData["FoodBevDinner"] = 0;
      ParsedSheetData["AlcoholLunch"] = 0;
      ParsedSheetData["AlcoholDinner"] = 0;
      ParsedSheetData["OnlineSales"] = 0;
      ParsedSheetData["LunchCovers"] = 0;
      ParsedSheetData["DinnerCovers"] = 0;
      ParsedSheetData["CashDeposit"] = 0;
      ParsedSheetData["PaidOut"] = 0;
      ParsedSheetData["86"] = 0;
      ParsedSheetData["CanceledOrder"] = 0;
      ParsedSheetData["Training"] = 0;
      ParsedSheetData["ChangedMind"] = 0; 
      ParsedSheetData["ServerError"] = 0;
      ParsedSheetData["ManagerMeal"] = 0;
      ParsedSheetData["DrawerMeal"] = 0;
      ParsedSheetData["Donation"] = 0;
      ParsedSheetData["EmployeeOnShift"] = 0;
      ParsedSheetData["EmployeeOffShift"] = 0;
      ParsedSheetData["InHousePromo"] = 0;
      ParsedSheetData["CobbTeacher"] = 0;
      ParsedSheetData["ManagerOwner"] = 0;
      ParsedSheetData["GoodCustomer"] = 0;
      ParsedSheetData["FirePolice"] = 0;
    }

    private void GatherDiscounts()
    {
      var startingIndex = _categoryRowIndexes["DiscountSummary"];
      var endingIndex = _categoryRowIndexes["GuestTableInformationByPeriod"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = _rawSheetData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString()?.ToLower();

        switch (rowTitle)
        {
          case "cobb county teacher":
            var cobbCountyTeacher = (double)currentRow[5];
            ParsedSheetData["CobbTeacher"] = cobbCountyTeacher;
            break;
          case "employee":
            var employeeOnShift = (double)currentRow[5];
            ParsedSheetData["EmployeeOnShift"] = employeeOnShift;
            break;
          case "employee off shift":
            var employeeOffShift = (double)currentRow[5];
            ParsedSheetData["EmployeeOffShift"] = employeeOffShift;
            break;
          case "fire dept":
            var fireDept = (double)currentRow[5];
            AddToExistingValue("FirePolice", fireDept);
            break;
          case "good customer":
            var goodCustomer = (double)currentRow[5];
            ParsedSheetData["GoodCustomer"] = goodCustomer;
            break;
          case "in-house promo":
            var inHousePromo = (double)currentRow[5];
            ParsedSheetData["InHousePromo"] = inHousePromo;
            break;
          case "manager":
            var manager = (double)currentRow[5];
            AddToExistingValue("ManagerOwner", manager);
            break;
          case "police":
            var police = (double)currentRow[5];
            AddToExistingValue("FirePolice", police);
            break;
          case null:
            break;
        }
      }
    }

    private void GatherComps()
    {
      var startingIndex = _categoryRowIndexes["CompSummary"];
      var endingIndex = _categoryRowIndexes["DiscountSummary"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = _rawSheetData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString()?.ToLower();

        switch (rowTitle)
        {
          case "manager meal":
            var managerMeal = (double)currentRow[5];
            ParsedSheetData["ManagerMeal"] = managerMeal;
            break;
          case "drawer meal":
            var drawerMeal = (double)currentRow[5];
            ParsedSheetData["DrawerMeal"] = drawerMeal;
            break;
          case "donation":
            var donation = (double)currentRow[5];
            ParsedSheetData["Donation"] = donation;
            break;
          case null:
            break;
        }
      }
    }

    private void GatherVoids()
    {
      var startingIndex = _categoryRowIndexes["VoidSummary"];
      var endingIndex = _categoryRowIndexes["CompSummary"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = _rawSheetData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString()?.ToLower();

        switch (rowTitle)
        {
          case "changed mind":
            var changedMind = (double)currentRow[5];
            ParsedSheetData["ChangedMind"] = changedMind;
            break;
          case "86":
            var eightysix = (double)currentRow[5];
            ParsedSheetData["86"] = eightysix;
            break;
          case "canceled order":
            var canceledOrder = (double)currentRow[5];
            ParsedSheetData["CanceledOrder"] = canceledOrder;
            break;
          case "server error":
            var serverError = (double)currentRow[5];
            ParsedSheetData["ServerError"] = serverError;
            break;
          case "training":
            var training = (double)currentRow[5];
            ParsedSheetData["Training"] = training;
            break;
          case null:
            break;
        }
      }
    }

    private void GatherCashSection()
    {
      var startingIndex = _categoryRowIndexes["Deposit"];
      var endingIndex = _rawSheetData.Rows.Count;

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = _rawSheetData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString()?.Trim().ToLower();

        switch (rowTitle)
        {
          case "net cash received":
            var cashDeposit = Convert.ToDouble(currentRow[7]);
            ParsedSheetData["CashDeposit"] = cashDeposit;
            break;
          case "paid out":
            var paidOut = Convert.ToDouble(currentRow[4]);
            ParsedSheetData["PaidOut"] = paidOut;
            break;
          case null:
            break;
        }
      }
    }

    private void GatherCovers()
    {
      var startingIndex = _categoryRowIndexes["GuestTableInformationByPeriod"];
      var endingIndex = _categoryRowIndexes["GuestTableInformationByProfitCenter"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = _rawSheetData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString()?.ToLower();

        var coversRowFound = rowTitle == "number guest served";
        if (!coversRowFound) continue;

        var lunchCovers = (double)currentRow[1];
        ParsedSheetData["LunchCovers"] = lunchCovers;

        var dinnerCovers = (double)currentRow[2];
        ParsedSheetData["DinnerCovers"] = dinnerCovers;
      }
    }

    private void GatherOnlineSalesValues()
    {
      var startingIndex = _categoryRowIndexes["PaymentSummary"];
      var endingIndex = _categoryRowIndexes["SummaryByPeriod"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = _rawSheetData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString()?.ToLower();

        switch (rowTitle)
        {
          case "grubhub":
            var grubhub = (double)currentRow[5];
            AddToExistingValue("OnlineSales", grubhub);
            break;
          case "doordash":
            var doordash = (double)currentRow[5];
            AddToExistingValue("OnlineSales", doordash);
            break;
          case "eatstreet":
            var eatstreet = (double)currentRow[5];
            AddToExistingValue("OnlineSales", eatstreet);
            break;
          case "slice":
            var slice = (double)currentRow[5];
            AddToExistingValue("OnlineSales", slice);
            break;
          case "menufy":
            var menufy = (double)currentRow[5];
            AddToExistingValue("OnlineSales", menufy);
            break;
          case null:
            break;
        }
      }
    }

    private void GatherSummaryByPeriodValues()
    {
      var startingIndex = _categoryRowIndexes["SummaryByPeriod"];
      var endingIndex = _categoryRowIndexes["SummaryByProfitCenter"];

      for (var i = startingIndex; i < endingIndex; i++)
      {
        var currentRow = _rawSheetData.Rows[i].ItemArray;
        var rowTitle = currentRow[0].ToString()?.ToLower();

        switch (rowTitle)
        {
          case "alcohol":
            var alcoholLunch = (double)currentRow[1];
            ParsedSheetData["AlcoholLunch"] = alcoholLunch;

            var alcoholDinner = (double)currentRow[2];
            ParsedSheetData["AlcoholDinner"] = alcoholDinner;
            break;
          case "beverage":
            var beverageLunch = (double)currentRow[1];
            AddToExistingValue("FoodBevLunch", beverageLunch);

            var beverageDinner = (double)currentRow[2];
            AddToExistingValue("FoodBevDinner", beverageDinner);
            break;
          case "delivery":
            var deliveryLunch = (double)currentRow[1];
            AddToExistingValue("FoodBevLunch", deliveryLunch);

            var deliveryDinner = (double)currentRow[2];
            AddToExistingValue("FoodBevDinner", deliveryDinner);
            break;
          case "food":
            var foodLunch = (double)currentRow[1];
            AddToExistingValue("FoodBevLunch", foodLunch);

            var foodDinner = (double)currentRow[2];
            AddToExistingValue("FoodBevDinner", foodDinner);
            break;
          case null:
            break;
        }
      }
    }

    private Dictionary<string, int> FindCategoryRowIndexes()
    {
      var categoryRowIndexes = new Dictionary<string, int>();

      for (var i = 0; i < _rawSheetData.Rows.Count; i++)
      {
        foreach (var item in _rawSheetData.Rows[i].ItemArray)
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

    private static DataTable IngestRawDataFromExcelSheet(string fileName)
    {
      Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
      var stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
      var reader = ExcelReaderFactory.CreateReader(stream);
      var dataSet = reader.AsDataSet();
      var rawSheetData = dataSet.Tables[0];
      return rawSheetData;
    }

    private void AddToExistingValue(string key, double value)
    {
      var newValue = value;
      var keyExists = ParsedSheetData.ContainsKey(key);
      if (keyExists)
      {
        var existingValue = ParsedSheetData[key];
        newValue = existingValue + value;
      }
      ParsedSheetData[key] = newValue;
    }
  }
}
