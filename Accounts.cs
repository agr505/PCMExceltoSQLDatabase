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
    public class globalexcel
    {
        public static string AccountIdentifierExcel, AccountNameExcel, ChannelTypeExcel, ProductTypeExcel, AssetTypeExcel, ConsultantExcel, AccountStatusExcel,
                    AxysIDExcel, CustodianExcel, CustodianAccountNumberExcel, TaxStatusExcel, BlockGroupExcel, ContractTypeExcel, CompositeExcel;
        public static DateTime? InceptionDateExcel;
        public static DateTime? TerminationDateExcel;
    }

    [Table(Name = "Accounts")]//Name of SQL Table this Object should be mapped to
    public class Accounts
    {

        //Compares Accounts Object that contains Excel data to Accounts object from SQL database and returns true if they are different meaning that rows have been changed in Excel file
        //since last insert and returns false if the row has not been changed 
        public bool compare(Accounts obj)
            {

        
                if (this.AccountIdentifier != obj.AccountIdentifier)
                {
                    return true;
                }
                if (this.AccountName != obj.AccountName)
                {
                    return true;
                }
                if (this.AccountStatus != obj.AccountStatus)
                {
                    return true;
                }
                if (this.AssetType != obj.AssetType)
                {
                    return true;
                }
                if (this.AxysID != obj.AxysID)
                {
                    return true;
                }
                if (this.BlockGroup != obj.BlockGroup)
                {
                    return true;
                }
                if (this.ChannelType != obj.ChannelType)
                {
                    return true;
                }
                if (this.Composite != obj.Composite)
                {
                    return true;
                }
                if (this.Consultant != obj.Consultant)
                {
                    return true;
                }
                if (this.Custodian != obj.Custodian)
                {
                    return true;
                }
                if (this.CustodianAccountNumber != obj.CustodianAccountNumber)
                {
                    return true;
                }
                if (this.InceptionDate != obj.InceptionDate)
                {
                    return true;
                }
                if (this.TaxStatus != obj.TaxStatus)
                {
                    return true;
                }
                if (this._TerminationDate != obj._TerminationDate)
                {
                    return true;
                }
                if (this._ContractType != obj._ContractType)
                {
                    return true;
                }
                if (this._ProductType != obj._ProductType)
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
            private string _AccountName;
            [Column(Storage = "_AccountName")]
            public string AccountName
            {
                get { return this._AccountName; }
                set { this._AccountName = value; }
            }
            private string _ChannelType;
            [Column(Storage = "_ChannelType")]
            public string ChannelType
            {
                get { return this._ChannelType; }
                set { this._ChannelType = value; }
            }
            private string _ProductType;
            [Column(Storage = "_ProductType")]
            public string ProductType
            {
                get { return this._ProductType; }
                set { this._ProductType = value; }
            }
            private string _AssetType;
            [Column(Storage = "_AssetType")]
            public string AssetType
            {
                get { return this._AssetType; }
                set { this._AssetType = value; }
            }
            private string _Consultant;
            [Column(Storage = "_Consultant")]
            public string Consultant
            {
                get { return this._Consultant; }
                set { this._Consultant = value; }
            }
            private string _AccountStatus;
            [Column(Storage = "_AccountStatus")]
            public string AccountStatus
            {
                get { return this._AccountStatus; }
                set { this._AccountStatus = value; }
            }
            private DateTime? _InceptionDate;
            [Column(Storage = "_InceptionDate")]
            public DateTime? InceptionDate
            {
                get { return this._InceptionDate; }
                set { this._InceptionDate = value; }
            }
            private DateTime? _TerminationDate;
            [Column(Storage = "_TerminationDate")]
            public DateTime? TerminationDate
            {
                get { return this._TerminationDate; }
                set { this._TerminationDate = value; }
            }
            private string _AxysID;
            [Column(Storage = "_AxysID")]
            public string AxysID
            {
                get { return this._AxysID; }
                set { this._AxysID = value; }
            }
            private string _Custodian;
            [Column(Storage = "_Custodian")]
            public string Custodian
            {
                get { return this._Custodian; }
                set { this._Custodian = value; }
            }
            private string _CustodianAccountNumber;
            [Column(Storage = "_CustodianAccountNumber")]
            public string CustodianAccountNumber
            {
                get { return this._CustodianAccountNumber; }
                set { this._CustodianAccountNumber = value; }
            }
            private string _TaxStatus;
            [Column(Storage = "_TaxStatus")]
            public string TaxStatus
            {
                get { return this._TaxStatus; }
                set { this._TaxStatus = value; }
            }
            private string _BlockGroup;
            [Column(Storage = "_BlockGroup")]
            public string BlockGroup
            {
                get { return this._BlockGroup; }
                set { this._BlockGroup = value; }
            }
            private string _ContractType;
            [Column(Storage = "_ContractType")]
            public string ContractType
            {
                get { return this._ContractType; }
                set { this._ContractType = value; }
            }
            private string _Composite;
            [Column(Storage = "_Composite")]
            public string Composite
            {
                get { return this._Composite; }
                set { this._Composite = value; }
            }
        

    }
}
