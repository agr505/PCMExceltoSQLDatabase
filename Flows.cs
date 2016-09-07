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
    public class globalexcel2
    {
        public static string AccountIdentifierExcel, TypeExcel, SubTypeExcel;
        public static int Id;
        public static decimal AmountExcel;
        public static DateTime? DateExcel;
    }
    [Table(Name = "Flows")]//Name of SQL Table this Object should be mapped to
    class Flows
    {



        //Object properties that will be mapped to SQL Table Columns
        private int _Id;
            [Column(IsPrimaryKey = true, Storage = "_Id")]
            public int Id
            {
                get { return this._Id; }
                set { this._Id = value; }
            }
            private string _AccountIdentifier;
            [Column(Storage = "_AccountIdentifier")]
            public string AccountIdentifier
            {
                get { return this._AccountIdentifier; }
                set { this._AccountIdentifier = value; }
            }
            private DateTime? _Date;
            [Column(IsPrimaryKey = true, Storage = "_Date")]
            public DateTime? Date
            {
                get { return this._Date; }
                set { this._Date = value; }
            }
            private decimal _Amount;
            [Column(Storage = "_Amount")]
            public decimal Amount
            {
                get { return this._Amount; }
                set { this._Amount = value; }
            }
            private string _Type;
            [Column(Storage = "_Type")]
            public string Type
            {
                get { return this._Type; }
                set { this._Type = value; }
            }
            private string _SubType;
            [Column(Storage = "_SubType")]
            public string SubType
            {
                get { return this._SubType; }
                set { this._SubType = value; }
            }

        
    }
}
