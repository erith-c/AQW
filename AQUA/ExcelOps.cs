#define EXCEL
#undef CONSOLE
#undef FILE

using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AQUA
{
    public partial class ExcelOps
    {
        static Excel.Application ExcelApp = new Excel.Application();
        static string FilePath = @"D:\jechase\erith-c\resources\test-resources\excel-test.xlsx";
        static string LogoPath = @"D:\jechase\erith-c\resources\icons-and-logos\APCLogoRight-notext.jpg";

        public enum Align
        {
            HLeft,
            HRight,
            HCenter,
            HJustify,
            VTop,
            VCenter,
            VBottom,
        }

        static Color APC_Blue = Color.FromArgb(0x003087);
        static Color Black = Color.FromArgb(0x000000);
        static Color Highlight = Color.FromArgb(0xffff00);
        static Color White = Color.FromArgb(0xffffff);
        static Color ShipGround = Color.FromArgb(0x00b050);
        static Color Ship2Day = Color.FromArgb(0x00b0f0);
        static Color ShipRed = Color.FromArgb(0xc00000);

        public static void InitializeWorkbook()
        {
            ExcelApp.Visible = true;
            Excel.Workbook WB = ExcelApp.Workbooks.Open(FilePath);
            QWSetup(WB);
            
        }

        public static void QWSetup(Workbook WB)
        {
            #region Setup
            // Create the Worksheet
            Worksheet QW = WB.Worksheets.Add();
            QW.Name = "Quotation Worksheet";
            QW.Activate();
            WB.Windows[1].Zoom = 130;

            // Set defaults
            QW.Cells.Font.Name = "Ebrima";
            QW.Cells.Font.Size = 10;

            // Sets the widths of columns A through X
            double[] ColumnWidths = new double[] { 
                4.43, 9.29, 9.29, 5.00, 7.86, 2.29,   // A : F
                4.57, 7.43, 5.00, 5.00, 5.86, 3.29,   // G : L
                4.00, 4.86, 5.14, 1.14, 1.14, 5.43,   // M : R
                1.29, 6.71, 6.71, 4.00, 4.00, 4.00    // S : X
            };

            // TODO: Set ActiveRange scope to be determined by data source
            Range ActiveRange_Start = QW.Range["A1"];
            Range ActiveRange_End = QW.Range["X42"];

            // Header <- Variable heights, # of rows, used on multiple pages
            double[] HeaderHeights = new double[]
            {
                22.50, 15.00, 12.75, 18.00, 6.00, 18.00, 6.00
            };

            // Markup <- Variable heights, # of rows, used on first page only
            double[] MarkupHeights = new double[]
            {
                15.00, 14.25, 14.25, 14.25, 14.25, 14.25, //  8 : 13
                14.25, 14.25, 14.25, 14.25, 14.25, 14.25, // 14 : 19
                14.25, 14.25, 14.25, 14.25, 6.00          // 20 : 24
            };

            // Materials List Header <- Set height, 1 row, used on first page only
            // Materials List Column Headers <- Set heights, 2 rows, used on multiple pages
            double[] MaterialsListHeights = new double[] { 18.00, 14.25, 15.75 };

            // Materials List <- static height @ 14.25pts, dynamic # of rows, expanded to multiple pages if necessary
            double MaterialsList_StdHeight = 14.25;

            Range QWActiveRange = QW.Range["A1", "X42"];

            SetColumnWidths(QWActiveRange, ColumnWidths);
            SetHeaderRows(QWActiveRange, HeaderHeights);
            SetMarkupRows(QWActiveRange, MarkupHeights, 8);
            SetMaterialsListHeaders(QWActiveRange, MaterialsListHeights, 25);
            SetMaterialsList(QWActiveRange, 15, 28, MaterialsList_StdHeight);
            #endregion Setup

            #region Header Object Declarations
            // Set Header Objects
            Shape APC_Logo = QW.Shapes.AddPicture(LogoPath,
                                                  Microsoft.Office.Core.MsoTriState.msoFalse,
                                                  Microsoft.Office.Core.MsoTriState.msoTrue,
                                                  0.00F, 0.00F, 227.52F, 54F);
            
            Range Title       = QW.Range["I1", "O1"];
            Range Version     = QW.Range["I2", "O2"];
            Range CfdBanner   = QW.Range["S1", "X1"];
            Range ApproverLbl = QW.Range["U2", "V2"];
            Range ApproverTxt = QW.Range["W2", "X2"];
            Range PageLbl     = QW.Range["U4"];
            Range PageTxt     = QW.Range["V4", "X4"];
            Range JobNameLbl  = QW.Range["A6", "B6"];
            Range JobNameTxt  = QW.Range["C6", "G6"];
            Range JobNumLbl   = QW.Range["H6", "I6"];
            Range JobNumTxt   = QW.Range["J6", "K6"];
            Range SalesLbl    = QW.Range["L6"];
            Range SalesTxt    = QW.Range["M6", "N6"];
            Range PMLbl       = QW.Range["O6"];
            Range PMTxt       = QW.Range["P6", "R6"];
            Range DateLbl     = QW.Range["S6", "T6"];
            Range DateTxt     = QW.Range["U6", "X6"];

            #endregion Header Object Declarations

            #region Header Formatting
            // Format Header Objects
            APC_Logo.PictureFormat.Contrast = 0.6F;
            
            #region Title
            
            Title.Merge();
            Title.Font.Size = 9;
            Title.Font.Bold = true;
            Title.Font.Italic = true;
            AlignRange(Title, V: Align.VCenter);

            // TODO: Set value to be dynamically adjusted based on data source.
            Title.Value = "Commercial Quotation Worksheet {C#}";
            #endregion Title
            
            #region Version
            Version.Merge();
            Version.Font.Superscript = true;
            AlignRange(Version, V: Align.VTop);

            // TODO: Set value to be dynamically adjusted based on data source.
            Version.Value = "Version: 0.0.1";
            #endregion Version

            #region Cfd Banner
            CfdBanner.Merge();
            CfdBanner.Borders.Weight = XlBorderWeight.xlMedium;
            CfdBanner.Borders.Color = Black;
            CfdBanner.Font.Bold = true;
            CfdBanner.Font.Color = White;
            CfdBanner.Interior.Color = APC_Blue;
            AlignRange(CfdBanner, V: Align.VCenter);
            CfdBanner.Value = "CONFIDENTIAL DOCUMENT";
            #endregion Cfd Banner

            #region Approver
            ApproverLbl.Merge();
            ApproverLbl.Font.Size = 8;
            ApproverLbl.Font.Bold = true;
            AlignRange(ApproverLbl, H: Align.HRight);
            ApproverLbl.Value = "Approved By:";

            ApproverTxt.Merge();
            ApproverTxt.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            AlignRange(ApproverTxt);
            #endregion Approver

            #region Page Number
            PageLbl.Font.Size = 8;
            PageLbl.Font.Bold = true;
            PageLbl.HorizontalAlignment = XlHAlign.xlHAlignRight;
            PageLbl.Value = "Page";

            PageTxt.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            PageTxt.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            PageTxt[1, 1].Value = 1;
            PageTxt[1, 2].Font.Size = 8;
            PageTxt[1, 2].Font.Bold = true;
            PageTxt[1, 2].Value = "of";

            // TODO: Dynamically set Page # based on data source.
            PageTxt[1, 3].Value = 1;

            #endregion Page Number

            #region Job Bar
            JobNameLbl.Merge();
            JobNameLbl.Font.Bold = true;
            AlignRange(JobNameLbl, H: Align.HRight);
            JobNameLbl.Value = "JOB NAME:";

            JobNameTxt.Merge();
            JobNameTxt.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            AlignRange(JobNameTxt);

            JobNumLbl.Merge();
            JobNumLbl.Font.Bold = true;
            AlignRange(JobNumLbl, H: Align.HRight);
            JobNumLbl.Value = "JOB NO:";

            JobNumTxt.Merge();
            JobNumTxt.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            AlignRange(JobNumTxt);

            SalesLbl.Merge();
            SalesLbl.Font.Bold = true;
            AlignRange(SalesLbl, H: Align.HRight);
            SalesLbl.Value = "SP:";

            SalesTxt.Merge();
            SalesTxt.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            AlignRange(SalesTxt);

            PMLbl.Merge();
            PMLbl.Font.Bold = true;
            AlignRange(PMLbl, H: Align.HRight);
            PMLbl.Value = "PM:";

            PMTxt.Merge();
            PMTxt.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            AlignRange(PMTxt);

            DateLbl.Merge();
            DateLbl.Font.Bold = true;
            AlignRange(DateLbl, H: Align.HRight);
            DateLbl.Value = "DATE:";

            DateTxt.Merge();
            DateTxt.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            AlignRange(DateTxt);
            #endregion Job Bar

            #endregion Header Formatting

            #region Markup Formatting
            #endregion Markup Formatting

            #region Materials List Formatting
            #endregion Materials List Formatting
        }
        # region Set Column Widths
        public static void SetColumnWidths(Range range, double[] widths_array)
        {
            try
            {
                for (int index = 0; (index < widths_array.Length); index++)
                {
                    range[1, index + 1].ColumnWidth = widths_array[index];
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"[ Error caught: Type { e.GetType() } ]\n>  { e.Message }");
            }
        }
        #endregion Set Column Widths
        #region Set Row Heights
        public static void SetHeaderRows(Range range, double[] heights_array)
        {
            try
            {
                for (int index = 0; index < heights_array.Length; index++)
                {
                    range[index + 1, 1].RowHeight = heights_array[index];
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"[ Error caught: Type { e.GetType() } ]\n>  { e.Message }");
            }
        }

        public static void SetMarkupRows(Range range, double[] heights_array, int rows_start)
        {
            try
            {
                for (int index = 0; index < heights_array.Length; index++)
                {
                    range[index + rows_start, 1].RowHeight = heights_array[index];
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"[ Error caught: Type { e.GetType() } ]\n>  { e.Message }");
            }
        }

        public static void SetMaterialsListHeaders(Range range, double[] heights_array, int rows_start)
        {
            try
            {
                for (int index = 0; index < heights_array.Length; index++)
                {
                    range[index + rows_start, 1].RowHeight = heights_array[index];
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"[ Error caught: Type { e.GetType() } ]\n>  { e.Message }");
            }
        }

        public static void SetMaterialsList(Range range, int number_of_rows, int rows_start, double row_height)
        {
            try
            {
                for (int index = 0; index <= number_of_rows; index++)
                {
                    range[index + rows_start, 1].RowHeight = row_height;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"[ Error caught: Type { e.GetType() } ]\n>  { e.Message }");
            }
        }
        #endregion Set Row Heights

        public static void AlignRange(Range range, Align H = Align.HCenter, Align V = Align.VBottom)
        {
            switch (H)
            {
                case Align.HCenter:
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    break;
                case Align.HLeft:
                    range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    break;
                case Align.HRight:
                    range.HorizontalAlignment = XlHAlign.xlHAlignRight;
                    break;
                case Align.HJustify:
                    range.HorizontalAlignment = XlHAlign.xlHAlignJustify;
                    break;
                default:
                    Console.WriteLine("Whoops! Not a valid alignment!");
                    break;
            }
            switch (V)
            {
                case Align.VTop:
                    range.VerticalAlignment = XlVAlign.xlVAlignTop;
                    break;
                case Align.VCenter:
                    range.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    break;
                case Align.VBottom:
                    range.VerticalAlignment = XlVAlign.xlVAlignBottom;
                    break;
                default:
                    Console.WriteLine("Whoops! Not a valid alignment!");
                    break;
            }
        }
    }
}
