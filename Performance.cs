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
    public class globalexcel3
    {
        public static string AccountIdentifier, AccountName, Currency, Category;
        public static DateTime? FromDate, ToDate;
        public static decimal TargetWeight, TotalReturn, LocalReturn, NetReturn, BeginMarketValue, NetFlowAmount, WeightedFlow, EndValue;








    }

    [Table(Name = "Performance")]//Name of SQL Table this Object should be mapped to
    class Performance
    {
       

        //Object properties that will be mapped to SQL Table Columns
        private string _AccountIdentifier;
        [Column(IsPrimaryKey = true, Storage = "_AccountIdentifier")]
        public string AccountIdentifier
        {
            get { return this._AccountIdentifier; }
            set { this._AccountIdentifier = value; }
        }

        private DateTime? _FromDate;
        [Column(IsPrimaryKey = true, Storage = "_FromDate")]
        public DateTime? FromDate
        {
            get { return this._FromDate; }
            set { this._FromDate = value; }
        }

        private DateTime? _ToDate;
        [Column(IsPrimaryKey = true, Storage = "_ToDate")]
        public DateTime? ToDate
        {
            get { return this._ToDate; }
            set { this._ToDate = value; }
        }
        private string _Currency;
        [Column(IsPrimaryKey = true, Storage = "_Currency")]
        public string Currency
        {
            get { return this._Currency; }
            set { this._Currency = value; }
        }
        private string _Category;
        [Column(IsPrimaryKey = true, Storage = "_Category")]
        public string Category
        {
            get { return this._Category; }
            set { this._Category = value; }
        }
        private decimal _TargetWeight;
        [Column(Storage = "_TargetWeight")]
        public decimal TargetWeight
        {
            get { return this._TargetWeight; }
            set { this._TargetWeight = value; }
        }
        private decimal _TotalReturn;
        [Column(Storage = "_TotalReturn")]
        public decimal TotalReturn
        {
            get { return this._TotalReturn; }
            set { this._TotalReturn = value; }
        }
        private decimal _LocalReturn;
        [Column(Storage = "_LocalReturn")]
        public decimal LocalReturn
        {
            get { return this._LocalReturn; }
            set { this._LocalReturn = value; }
        }
        private decimal _NetReturn;
        [Column(Storage = "_NetReturn")]
        public decimal NetReturn
        {
            get { return this._NetReturn; }
            set { this._NetReturn = value; }
        }
        private decimal _BeginMarketValue;
        [Column(Storage = "_BeginMarketValue")]
        public decimal BeginMarketValue
        {
            get { return this._BeginMarketValue; }
            set { this._BeginMarketValue = value; }
        }
        private decimal _NetFlowAmount;
        [Column(Storage = "_NetFlowAmount")]
        public decimal NetFlowAmount
        {
            get { return this._NetFlowAmount; }
            set { this._NetFlowAmount = value; }
        }
        private decimal _WeightedFlowsAmount;
        [Column(Storage = "_WeightedFlowsAmount")]
        public decimal WeightedFlowsAmount
        {
            get { return this._WeightedFlowsAmount; }
            set { this._WeightedFlowsAmount = value; }
        }
        private decimal _EndValue;
        [Column(Storage = "_EndValue")]
        public decimal EndValue
        {
            get { return this._EndValue; }
            set { this._EndValue = value; }
        }









    }
}
