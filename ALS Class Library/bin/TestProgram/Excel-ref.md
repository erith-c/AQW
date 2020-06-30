# C# Excel Reference #

``` C#
static void DisplayInExcel(IEnumerable<Account> accounts)
{
    var excelApp = new Excel.Application();
    excelApp.Visible = true;

    excelApp.Workbooks.Add();

    Excel._Worksheet workSheet = excelApp.ActiveSheet;

    workSheet.Cells[1, "A"] = "ID Number";
    workSheet.Cells[1, "B"] = "Current Balance";

    var row = 1;
    foreach (var acct in accounts)
    {
        row++;
        workSheet.Cells[row, "A"] = acct.ID;
        workSheet.Cells[row, "B"] = acct.Balance;
    }

    workSheet.Range["A1", "B3"].AutoFormat(
        Format: Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
}
```