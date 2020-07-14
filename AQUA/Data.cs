#define CONSOLE
#define EXCEL
#define FILE

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

[module: Author("Erith", version = "0.0.1")]

namespace ALS.AQUA
{
    #region Data Class
    public class Data
    {
        public abstract class Base<T>
        {
            public abstract string ID { get; }

            public abstract void AddRecord(T item);
            public abstract T GetRecord(string name);
            public abstract void DeleteRecord(T item);
            public abstract void PrintRecords();
        }
        public abstract class Record
        {
            public abstract string ID { get; }
        }
    }
    #endregion
    #region Inventory Table
    public class InventoryTable : Data.Base<InventoryItem>
    {
        public override string ID { get; }

        public override void AddRecord(InventoryItem item)
        {
            throw new NotImplementedException();
        }
        public override InventoryItem GetRecord(string name)
        {
            throw new NotImplementedException();
        }
        public override void DeleteRecord(InventoryItem item)
        {
            throw new NotImplementedException();
        }
        public override void PrintRecords()
        {
            throw new NotImplementedException();
        }
    }
    #endregion
    #region Inventory Item
    public class InventoryItem : Data.Record
    {
        public const int CNS_ValueLength = 30;
        // Properties
        public override string ID { get; }
        public static int ItemID { get; set; }
        public string PartID { get; set; }
        public string Description { get; set; }
        public double OrderPrice { get; set; }
        public double Quantity { get; set; }

        // Methods
        public override string ToString()
        {
#if CONSOLE
            return $"//================================================\\\\\n" +
                   $"|| ID:  {ID}  |  PartID:  {PartID,-(CNS_ValueLength - 10)}||\n" +
                   $"||------------------------------------------------||\n" +
                   $"|| > Description:   {'"' + Description + '"',-CNS_ValueLength}||\n" +
                   $"|| > Quantity:      {Quantity,-CNS_ValueLength}||\n" +
                   $"|| > OrderPrice:    $ {OrderPrice,-(CNS_ValueLength - 2):f2}||\n" +
                   $"\\\\================================================//\n";
#elif EXCEL
#elif FILE
#endif
        }

        // Constructors
        public InventoryItem ()
        {
            ++ItemID;
            this.ID = (Convert.ToInt32(ID, 16) + ItemID).ToString("X8");
        }
        public InventoryItem(string name)
        {
            this.PartID = name;
            ++ItemID;
            this.ID = (Convert.ToInt32(ID, 16) + ItemID).ToString("X8");
        }
    }
    #endregion
    #region QW
    namespace QW
    {
        #region Material Line Item
        public class MaterialLineItem : ALS.AQUA.InventoryItem
        {
            // Properties
            public bool ItemSource { get; set; }
            public DateTime OrderDate { get; set; }
            public DateTime ArrivalDate { get; set; }
            public new double Quantity { get; set; }
            public double OrderQuantity { get; set; }
            public double SellPrice { get; set; }
            public string PO_ID { get; set; }

            // Methods
            public override string ToString()
            {
                if (ItemSource == true)
                {
#if CONSOLE
                    return $"//================================================\\\\\n" +
                           $"|| ID:  {ID}  |  PartID:  {PartID,-(CNS_ValueLength - 10)}||\n" +
                           $"||------------------------------------------------||\n" +
                           $"|| > Description:   {'"' + Description + '"', -CNS_ValueLength}||\n" +
                           $"|| > Item Source:   {"<Ordered>", -CNS_ValueLength}||\n" +
                           $"|| > Quantity:      {Quantity, -CNS_ValueLength}||\n" +
                           $"|| > Sell Price:    $ {SellPrice, -(CNS_ValueLength - 2):f2}||\n" +
                           $"|| > Order Date:    {OrderDate, -CNS_ValueLength:d}||\n" +
                           $"|| > Arrival Date:  {ArrivalDate, -CNS_ValueLength:d}||\n" +
                           $"|| > Order Qty:     {OrderQuantity, -CNS_ValueLength}||\n" +
                           $"|| > Order Price:   $ {OrderPrice, -(CNS_ValueLength - 2):f2}||\n" +
                           $"|| > PO Number:     {PO_ID, -CNS_ValueLength}||\n" +
                           $"\\\\================================================//\n";
#elif EXCEL
                    // TODO: Return Excel-formatted output
#elif FILE
                    // TODO: Return WARDEN-formatted output
#endif
                }
                else
                {
#if CONSOLE
                    return $"//================================================\\\\\n" +
                           $"|| ID:  {ID}  |  PartID:  {PartID, -(CNS_ValueLength - 10)}||\n" +
                           $"||------------------------------------------------||\n" +
                           $"|| > Description:   {'"' + Description + '"',-CNS_ValueLength}||\n" +
                           $"|| > Quantity:      {Quantity,-CNS_ValueLength}||\n" +
                           $"|| > OrderPrice:    $ {OrderPrice,-(CNS_ValueLength - 2):f2}||\n" +
                           $"|| > Item Source:   {"<Stock>", -CNS_ValueLength}||\n" +
                           $"\\\\================================================//\n";
#elif EXCEL
#elif FILE
#endif
                }
            }

            // Constructors
            public MaterialLineItem()
                : base()
            {
                
            }
            public MaterialLineItem(string name)
                : base(name)
            {

            }
        }
        #endregion
    }
    #endregion
}
