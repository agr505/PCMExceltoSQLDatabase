using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Linq;
using System.Data.Linq.Mapping;
namespace PCMExceltoSQLDatabase
{
    //Class used for collecting Excel Data and then for inputting it into LINQ Entity object which is inserted into Database. It is used since it has all static
    //fields so it can be used in a global scope.
    public class globalexcel4
    {
        public static string AccountIdentifier;

        public static decimal TotalMarketValue, CashBalance;

    }
    [Table(Name = "Assets")]//Name of SQL Table this Object should be mapped to
    class Assets
    {
        //Compares Assets Object that contains Excel data to Assets object from SQL database and returns true if they are different meaning that rows have been changed in Excel file
        //since last insert and returns false if the row has not been changed 
        public bool compare(Assets obj) 
        {


            if (this.AccountIdentifier != obj.AccountIdentifier)
            {
                return true;
            }
            if (this.TotalMarketValue != obj.TotalMarketValue)
            {
                return true;
            }
            if (this.CashBalance != obj.CashBalance)
            {
                return true;
            }
          
            return false;

        }
        //Object properties that will be mapped to SQL Table Columns
        private string _AccountIdentifier;
        [Column(IsPrimaryKey = true, Storage = "_AccountIdentifier")]
        public string AccountIdentifier
        {
            get { return this._AccountIdentifier; }
            set { this._AccountIdentifier = value; }
        }

        private decimal _TotalMarketValue;
        [Column(Storage = "_TotalMarketValue")]
        public decimal TotalMarketValue
        {
            get { return this._TotalMarketValue; }
            set { this._TotalMarketValue = value; }
        }
        private decimal _CashBalance;
        [Column(Storage = "_CashBalance")]
        public decimal CashBalance
        {
            get { return this._CashBalance; }
            set { this._CashBalance = value; }
        }
    }
}
