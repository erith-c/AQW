using System;
using Excel = Microsoft.Office.Interop.Excel;
using APC = ALS.APC;

namespace TestProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            var record = new APC.DataRecord();
            Console.WriteLine($"Record: {record.Name}\n" +
                $"Worksheet: {record.Source}\n" +
                $"Cell: {record.CellRef}\n" +
                $"Value: {record.Value}");
            /*var bankAccounts = new List<Account>
            {
                new Account
                {
                    ID = 345678,
                    Balance = 541.27
                },
                new Account
                {
                    ID = 1230221,
                    Balance = -127.44
                }
            };
            DisplayInExcel(bankAccounts);
        }

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
                Format: Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);*/
        }
    }/*
    public class Account
    {
        public int ID { get; set; }
        public double Balance { get; set; }
    }*/
}
