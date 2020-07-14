#define CONSOLE
//#define EXCEL
//#define FILE

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ALS.AQUA;

public static class Test
{
    public static void Run()
    {
        try
        {
            // Represents an item in the APC Catalog.
            InventoryItem item1 = new InventoryItem("TESTPART1");
            Console.WriteLine(item1.ToString());

            // Represents an item on a QW that is in stock.
            ALS.AQUA.QW.MaterialLineItem item2 = new ALS.AQUA.QW.MaterialLineItem("TESTPART2");
            Console.WriteLine(item2.ToString());

            // Represents an item on a QW that is ordered.
            ALS.AQUA.QW.MaterialLineItem item3 = new ALS.AQUA.QW.MaterialLineItem("HU-920PMNNEKMA003");
            item3.ItemSource = true;
            Console.WriteLine(item3.ToString());
        }
        catch (Exception e)
        {
            Console.WriteLine($"Caught exception: {e.ToString()}");
        }
    }
}