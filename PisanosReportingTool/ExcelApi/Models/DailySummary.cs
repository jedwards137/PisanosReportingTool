using System;

namespace ExcelApi.Models
{
  public class DailySummary
  {
    public DateTime Date { get; set; }
    public SalesComparison SalesComparison { get; set; }

    public Covers Covers { get; set; }

    public Cash Cash { get; set; }

    public FoodVoids FoodVoids { get; set; }

    public FoodComps FoodComps { get; set; }

    public FoodDiscounts FoodDiscounts { get; set; }
  }
}
