using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Linq;
using System.Data.Linq;
using System.Data.Linq.Mapping;
using Microsoft.Office.Interop.Excel;
using System.Linq.Expressions;
using System.Data.Linq.SqlClient;
using System.Collections.Generic;
using System.Windows.Forms;
using static PCMExceltoSQLDatabase.Accounts;
using static PCMExceltoSQLDatabase.Performance;
namespace PCMExceltoSQLDatabase
{
    public partial class PCMExceltoSQLDatabase : Form
    {
        public string exfile; //Excel file name and path
        private bool _tab1confirmation;
        public bool tab1confirmation
        {
            get { return this._tab1confirmation; }
            set { this._tab1confirmation = value; }
        }
        private string _database;

        public string database //database connection string
        {
            get { return this._database; }
            set { this._database = value; }
        }
        private string _tablename;

        public string tablename
        {
            get { return this._tablename; }
            set { this._tablename = value; }
        }
        public PCMExceltoSQLDatabase()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.tablename = textBox2.Text;

            DataContext db = new DataContext(@"*rest of connection string" + textBox1.Text + "*rest of connection string" + Username.Text + "*rest of connection string" + Password.Text + "*rest of connection string");
            bool exist = db.DatabaseExists();



            if (exist)
            {
                this.printtext.Text = "The Database you selected exists";

                this.database = @"*rest of connection string" + textBox1.Text + "*rest of connection string" + Username.Text + "*rest of connection string" + Password.Text + "*rest of connection string";
                if (this.tablename != null)
                {
                    this.tab1confirmation = true;

                }


            }
            else
            {
                this.printtext.Text = "The Database you selected does not exist or your Username or Password is incorrect";

                this.tab1confirmation = false;

            }
        }

        private void OpenInsertUpdateTable(object sender, EventArgs e)
        {
            if (this.tab1confirmation == true && this.InsertUpdateTable.SelectedTab == tabPage2)
            {
                this.label5.Text = "Proceed and fill out the needed information.";
                this.Load += new System.EventHandler(this.Form1_Load);
            }
            else
            {
                this.label5.Text = "Complete Database and Excel File Section before proceeding!";
                this.Load += new System.EventHandler(this.Form1_Load);
            }
        }




        private void button1_Click(object sender, EventArgs e) //Excel file selection 
        {
            int size = -1;
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                exfile = openFileDialog1.FileName;


            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Insertsubmit_Click(object sender, EventArgs e)
        {
            int startrow = int.Parse(StartingRow.Text); //Assigns user designated starting row for the index to begin Excel row iteration at
            int endrow = int.Parse(EndRow.Text); //Assigns user designated ending row for the index to end Excel row iteration at

            //If and Else statements used to call specific fucntion corrensponding to the selected table and if the user selected to Insert or Update
            if (Accountscheck.Checked)
            {
                if (Insert.Checked)
                {
                    AccountsInsert(startrow, endrow);
                }
                else if (Update.Checked)
                {
                    AccountsUpdate(startrow, endrow);
                }
            }
            else if (Flowscheck.Checked)
            {
                FlowsInsertUpdate(startrow, endrow);

            }
            else if (Peformancecheck.Checked)
            {
                if (Insert.Checked)
                {
                    PerformanceInsert(startrow, endrow);
                }
                else if (Update.Checked)
                {
                    PerformanceUpdate(startrow, endrow);
                }
            }
            else if (Assetscheck.Checked)
            {
                if (Insert.Checked)
                {
                    AssetsInsert(startrow, endrow);
                }
                else if (Update.Checked)
                {
                    AssetsUpdate(startrow, endrow);
                }
            }
            else if (Activitycheck.Checked)
            {
                ActivityInsertUpdate(startrow, endrow);
            }

        }


        private void AccountsInsert(int startrow, int endrow)
        {
            this.Progress.Text = "Working on it...";//update status of insert
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status

            DataContext db = new DataContext(this.database); //created and initializes DataContext variable using the database connection string
            try
            {
                //create and initialize Excel application, workbook, and worksheet variables
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Workbook wkb = null;
                Worksheet sheet = null;

                wkb = excel.Workbooks.Open(exfile,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing);//open Excel file


                sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;//assigns specified worksheet to worksheet variable


                Table<Accounts> accounts = db.GetTable<Accounts>();//creates and initializes LINQ entity table object with table from SQL database

                Range range;


                for (int i = startrow, j = 1; i < endrow; j++) //starts at startrow and ends at endrow populating global excel object based on corresponding column numbers
                {

                    range = sheet.Cells[i, j];


                    if (j == 1)
                    {
                        globalexcel.AccountIdentifierExcel = range.Text.ToString();
                    }
                    else if (j == 2)
                    {
                        globalexcel.AccountNameExcel = range.Text.ToString();
                    }
                    else if (j == 3)
                    {
                        globalexcel.ChannelTypeExcel = range.Text.ToString();
                    }
                    else if (j == 4)
                    {
                        globalexcel.ProductTypeExcel = range.Text.ToString();
                    }
                    else if (j == 5)
                    {
                        globalexcel.AssetTypeExcel = range.Text.ToString();
                    }
                    else if (j == 6)
                    {
                        globalexcel.ConsultantExcel = range.Text.ToString();
                    }

                    else if (j == 7)
                    {
                        globalexcel.AccountStatusExcel = range.Text.ToString();
                    }
                    else if (j == 8)
                    {

                        if (range.Text.ToString() == "")
                        {
                            globalexcel.InceptionDateExcel = null;
                        }
                        else
                        {
                            globalexcel.InceptionDateExcel = DateTime.Parse(range.Text.ToString());
                        }
                    }
                    else if (j == 9)
                    {
                        if (range.Text.ToString() == "")
                        {
                            globalexcel.TerminationDateExcel = null;
                        }
                        else
                        {
                            globalexcel.TerminationDateExcel = DateTime.Parse(range.Text.ToString());
                        }

                    }
                    else if (j == 10)
                    {
                        globalexcel.AxysIDExcel = range.Text.ToString();
                    }
                    else if (j == 11)
                    {
                        globalexcel.CustodianExcel = range.Text.ToString();
                    }
                    else if (j == 12)
                    {
                        globalexcel.CustodianAccountNumberExcel = range.Text.ToString();
                    }
                    else if (j == 13)
                    {
                        globalexcel.TaxStatusExcel = range.Text.ToString();
                    }
                    else if (j == 14)
                    {
                        globalexcel.BlockGroupExcel = range.Text.ToString();
                    }
                    else if (j == 15)
                    {
                        globalexcel.ContractTypeExcel = range.Text.ToString();
                    }
                    else if (j == 16)
                    {
                        globalexcel.CompositeExcel = range.Text.ToString();
                        accounts.InsertOnSubmit( // creates and inserts new Accounts object which is one row in the Accounts SQL Table
                                                      new Accounts
                                                      {
                                                          AccountIdentifier = globalexcel.AccountIdentifierExcel,
                                                          AccountName = globalexcel.AccountNameExcel,
                                                          ChannelType = globalexcel.ChannelTypeExcel,
                                                          ProductType = globalexcel.ProductTypeExcel,
                                                          AssetType = globalexcel.AssetTypeExcel,
                                                          Consultant = globalexcel.ConsultantExcel,
                                                          AccountStatus = globalexcel.AccountStatusExcel,
                                                          InceptionDate = globalexcel.InceptionDateExcel,
                                                          TerminationDate = globalexcel.TerminationDateExcel,
                                                          AxysID = globalexcel.AxysIDExcel,
                                                          Custodian = globalexcel.CustodianExcel,
                                                          CustodianAccountNumber = globalexcel.CustodianAccountNumberExcel,
                                                          TaxStatus = globalexcel.TaxStatusExcel,
                                                          BlockGroup = globalexcel.BlockGroupExcel,
                                                          ContractType = globalexcel.ContractTypeExcel,
                                                          Composite = globalexcel.CompositeExcel,
                                                      });

                        db.SubmitChanges();
                        j = 0;
                        i++;





                    }







                }
                this.Progress.Text = "Finished";
                this.Load += new System.EventHandler(this.Form1_Load);

            }
            catch (Exception ex)
            {
                //if you need to handle stuff
                Console.WriteLine(ex);
            }
        }
        private void AccountsUpdate(int startrow, int endrow)
        {
            this.Progress.Text = "Working on it...";//update status of update
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status

            DataContext db = new DataContext(this.database);//created and initializes DataContext variable using the database connection string

            //create and initialize Excel application, workbook, and worksheet variables
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = null;
            Worksheet sheet = null;

            wkb = excel.Workbooks.Open(exfile,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing);//open Excel file


            sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;//assigns specified worksheet to worksheet variable


            Table<Accounts> accounts = db.GetTable<Accounts>();//creates and initializes LINQ entity table object with table from SQL database
            bool different = false;
            Range range;



            Accounts[] E = new Accounts[endrow];//creates an array of Accounts objects to represent the rows in the Excel file






            for (int i = startrow, j = 1; i < endrow; j++)//starts at startrow and ends at endrow populating global excel object based on corresponding column numbers
            {

                range = sheet.Cells[i, j];




                if (j == 1)
                {
                    globalexcel.AccountIdentifierExcel = range.Text.ToString();
                }
                else if (j == 2)
                {
                    globalexcel.AccountNameExcel = range.Text.ToString();
                }
                else if (j == 3)
                {
                    globalexcel.ChannelTypeExcel = range.Text.ToString();
                }
                else if (j == 4)
                {
                    globalexcel.ProductTypeExcel = range.Text.ToString();
                }
                else if (j == 5)
                {
                    globalexcel.AssetTypeExcel = range.Text.ToString();
                }
                else if (j == 6)
                {
                    globalexcel.ConsultantExcel = range.Text.ToString();
                }

                else if (j == 7)
                {
                    globalexcel.AccountStatusExcel = range.Text.ToString();
                }
                else if (j == 8)
                {

                    if (range.Text.ToString() == "")
                    {
                        globalexcel.InceptionDateExcel = null;
                    }
                    else
                    {
                        globalexcel.InceptionDateExcel = DateTime.Parse(range.Text.ToString());
                    }
                }
                else if (j == 9)
                {
                    if (range.Text.ToString() == "")
                    {
                        globalexcel.TerminationDateExcel = null;
                    }
                    else
                    {
                        globalexcel.TerminationDateExcel = DateTime.Parse(range.Text.ToString());
                    }

                }
                else if (j == 10)
                {
                    globalexcel.AxysIDExcel = range.Text.ToString();
                }
                else if (j == 11)
                {
                    globalexcel.CustodianExcel = range.Text.ToString();
                }
                else if (j == 12)
                {
                    globalexcel.CustodianAccountNumberExcel = range.Text.ToString();
                }
                else if (j == 13)
                {
                    globalexcel.TaxStatusExcel = range.Text.ToString();
                }
                else if (j == 14)
                {
                    globalexcel.BlockGroupExcel = range.Text.ToString();
                }
                else if (j == 15)
                {
                    globalexcel.ContractTypeExcel = range.Text.ToString();
                }
                else if (j == 16)
                {
                    globalexcel.CompositeExcel = range.Text.ToString();
                    E[i] = new Accounts
                    { // creates a new Accounts object which represents a row in the Excel file and stores it in the Excel row object array for later comparision with SQL Table data
                        AccountIdentifier = globalexcel.AccountIdentifierExcel,
                        AccountName = globalexcel.AccountNameExcel,
                        ChannelType = globalexcel.ChannelTypeExcel,
                        ProductType = globalexcel.ProductTypeExcel,
                        AssetType = globalexcel.AssetTypeExcel,
                        Consultant = globalexcel.ConsultantExcel,
                        AccountStatus = globalexcel.AccountStatusExcel,
                        InceptionDate = globalexcel.InceptionDateExcel,
                        TerminationDate = globalexcel.TerminationDateExcel,
                        AxysID = globalexcel.AxysIDExcel,
                        Custodian = globalexcel.CustodianExcel,
                        CustodianAccountNumber = globalexcel.CustodianAccountNumberExcel,
                        TaxStatus = globalexcel.TaxStatusExcel,
                        BlockGroup = globalexcel.BlockGroupExcel,
                        ContractType = globalexcel.ContractTypeExcel,
                        Composite = globalexcel.CompositeExcel,
                    };

                    j = 0;
                    i++;



                }







            }


            int index = startrow;

            foreach (Accounts row in accounts)
            {



                //Compares Accounts Object that contains Excel data to Accounts object from SQL database and returns true if they are different meaning that rows have been changed in Excel file
                //since last insert and returns false if the row has not been changed 
                different = E[index].compare(row);
                if (different)//if the row has been changed in the excel file then insert the updated Excel row and delete the SQL table row
                {
                    accounts.InsertOnSubmit(E[index]);
                    accounts.DeleteOnSubmit(row);
                    db.SubmitChanges();
                }
                index++;
            }
            if (index < endrow)
            {

                for (int j = 1; index < endrow; j++)//starts at startrow and ends at endrow populating global excel object based on corresponding column numbers
                {
                    range = sheet.Cells[index, j];




                    if (j == 1)
                    {
                        globalexcel.AccountIdentifierExcel = range.Text.ToString();
                    }
                    else if (j == 2)
                    {
                        globalexcel.AccountNameExcel = range.Text.ToString();
                    }
                    else if (j == 3)
                    {
                        globalexcel.ChannelTypeExcel = range.Text.ToString();
                    }
                    else if (j == 4)
                    {
                        globalexcel.ProductTypeExcel = range.Text.ToString();
                    }
                    else if (j == 5)
                    {
                        globalexcel.AssetTypeExcel = range.Text.ToString();
                    }
                    else if (j == 6)
                    {
                        globalexcel.ConsultantExcel = range.Text.ToString();
                    }

                    else if (j == 7)
                    {
                        globalexcel.AccountStatusExcel = range.Text.ToString();
                    }
                    else if (j == 8)
                    {

                        if (range.Text.ToString() == "")
                        {
                            globalexcel.InceptionDateExcel = null;
                        }
                        else
                        {
                            globalexcel.InceptionDateExcel = DateTime.Parse(range.Text.ToString());
                        }
                    }
                    else if (j == 9)
                    {
                        if (range.Text.ToString() == "")
                        {
                            globalexcel.TerminationDateExcel = null;
                        }
                        else
                        {
                            globalexcel.TerminationDateExcel = DateTime.Parse(range.Text.ToString());
                        }

                    }
                    else if (j == 10)
                    {
                        globalexcel.AxysIDExcel = range.Text.ToString();
                    }
                    else if (j == 11)
                    {
                        globalexcel.CustodianExcel = range.Text.ToString();
                    }
                    else if (j == 12)
                    {
                        globalexcel.CustodianAccountNumberExcel = range.Text.ToString();
                    }
                    else if (j == 13)
                    {
                        globalexcel.TaxStatusExcel = range.Text.ToString();
                    }
                    else if (j == 14)
                    {
                        globalexcel.BlockGroupExcel = range.Text.ToString();
                    }
                    else if (j == 15)
                    {
                        globalexcel.ContractTypeExcel = range.Text.ToString();
                    }
                    else if (j == 16)
                    {
                        globalexcel.CompositeExcel = range.Text.ToString();
                        accounts.InsertOnSubmit(new Accounts
                        { // creates and inserts new Accounts object which is one row in the Accounts SQL Table
                            AccountIdentifier = globalexcel.AccountIdentifierExcel,
                            AccountName = globalexcel.AccountNameExcel,
                            ChannelType = globalexcel.ChannelTypeExcel,
                            ProductType = globalexcel.ProductTypeExcel,
                            AssetType = globalexcel.AssetTypeExcel,
                            Consultant = globalexcel.ConsultantExcel,
                            AccountStatus = globalexcel.AccountStatusExcel,
                            InceptionDate = globalexcel.InceptionDateExcel,
                            TerminationDate = globalexcel.TerminationDateExcel,
                            AxysID = globalexcel.AxysIDExcel,
                            Custodian = globalexcel.CustodianExcel,
                            CustodianAccountNumber = globalexcel.CustodianAccountNumberExcel,
                            TaxStatus = globalexcel.TaxStatusExcel,
                            BlockGroup = globalexcel.BlockGroupExcel,
                            ContractType = globalexcel.ContractTypeExcel,
                            Composite = globalexcel.CompositeExcel,
                        });
                        db.SubmitChanges();
                        j = 0;
                        index++;



                    }

                }





            }
            this.Progress.Text = "Finished";//update status of update
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status
        }
        private void FlowsInsertUpdate(int startrow, int endrow)
        {

            this.Progress.Text = "Working on it...";//update status 
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status

            DataContext db = new DataContext(this.database);

            //create and initialize Excel application, workbook, and worksheet variables
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = null;
            Worksheet sheet = null;

            //Open Excel file
            wkb = excel.Workbooks.Open(exfile,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing);

            sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;//assigns specified worksheet to worksheet variable
            Table<Flows> flows = db.GetTable<Flows>();//creates and initializes LINQ entity table object with table from SQL database



            Range range;
            int nextrow = 0;
            globalexcel2.Id = 0;

            if (Update.Checked)
            { //in order to update, the last row in the Flows SQL Table is queried and its ID is used to locate the location in the Excel file to begin updating from
                IEnumerable<Flows> query = flows;

                Flows result = query.Last<Flows>();
                globalexcel2.Id = result.Id + 1;
                nextrow = result.Id + 1;


            }
            else
            {
                //if Update was not selected then insert will begin from the specified starting row
                nextrow = startrow;
            }

            for (int j = 1; nextrow < endrow + 1; j++)//starts at startrow and ends at endrow populating global excel object based on corresponding column numbers
            {




                range = sheet.Cells[nextrow, j];





                if (j == 1)
                {
                    globalexcel2.AccountIdentifierExcel = range.Text.ToString();
                }
                else if (j == 2)
                {
                    globalexcel2.DateExcel = DateTime.Parse(range.Text.ToString());
                }
                else if (j == 3)
                {
                    globalexcel2.AmountExcel = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 4)
                {
                    globalexcel2.TypeExcel = range.Text.ToString();
                }

                else if (j == 5)
                {
                    globalexcel2.Id++;
                    globalexcel2.SubTypeExcel = range.Text.ToString();
                    flows.InsertOnSubmit(
                                                       new Flows
                                                       { // creates and inserts new Flows object which is one row in the Accounts SQL Table
                                                           Id = globalexcel2.Id,
                                                           AccountIdentifier = globalexcel2.AccountIdentifierExcel,
                                                           Date = globalexcel2.DateExcel,
                                                           Amount = globalexcel2.AmountExcel,
                                                           Type = globalexcel2.TypeExcel,
                                                           SubType = globalexcel2.SubTypeExcel,

                                                       });

                    db.SubmitChanges();
                    j = 0;
                    nextrow++;

                }

            }
            this.Progress.Text = "Finished";//update status 
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status

        }
        private void PerformanceInsert(int startrow, int endrow)
        {
            this.Progress.Text = "Working on it...";
            this.Load += new System.EventHandler(this.Form1_Load);

            DataContext db = new DataContext(this.database);

            //create and initialize Excel application, workbook, and worksheet variables
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = null;
            Worksheet sheet = null;

            wkb = excel.Workbooks.Open(exfile,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing);//open Excel file


            sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;//assigns specified worksheet to worksheet variable
            Table<Performance> perf = db.GetTable<Performance>();//creates and initializes LINQ entity table object with table from SQL database



            Range range;



            for (int j = 1, nextrow = startrow; nextrow < endrow + 1; j++)//starts at startrow and ends at endrow populating global excel object based on corresponding column numbers
            {

                range = sheet.Cells[nextrow, j];




                if (j == 1)
                {
                    globalexcel3.AccountIdentifier = range.Text.ToString();
                }
                else if (j == 2)
                {

                    globalexcel3.FromDate = DateTime.Parse(range.Text.ToString());
                }
                else if (j == 3)
                {
                    globalexcel3.ToDate = DateTime.Parse(range.Text.ToString());
                }
                else if (j == 4)
                {
                    globalexcel3.Currency = range.Text.ToString();
                }
                else if (j == 5)
                {
                    globalexcel3.Category = range.Text.ToString();
                }
                else if (j == 6)
                {
                    globalexcel3.TargetWeight = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 7)
                {
                    globalexcel3.TotalReturn = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 8)
                {
                    globalexcel3.LocalReturn = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 9)
                {
                    globalexcel3.NetReturn = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 10)
                {
                    globalexcel3.BeginMarketValue = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 11)
                {
                    globalexcel3.NetFlowAmount = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 12)
                {
                    globalexcel3.WeightedFlow = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 13)
                {
                    globalexcel3.EndValue = System.Convert.ToDecimal(range.Value);



                    perf.InsertOnSubmit(
                                                   new Performance
                                                   { // creates and inserts new Performance object which is one row in the Performance SQL Table
                                                       AccountIdentifier = globalexcel3.AccountIdentifier,
                                                       FromDate = globalexcel3.FromDate,
                                                       ToDate = globalexcel3.ToDate,
                                                       Currency = globalexcel3.Currency,
                                                       Category = globalexcel3.Category,
                                                       TargetWeight = globalexcel3.TargetWeight,
                                                       TotalReturn = globalexcel3.TotalReturn,
                                                       LocalReturn = globalexcel3.LocalReturn,
                                                       NetReturn = globalexcel3.NetReturn,
                                                       BeginMarketValue = globalexcel3.BeginMarketValue,
                                                       NetFlowAmount = globalexcel3.NetFlowAmount,
                                                       WeightedFlowsAmount = globalexcel3.WeightedFlow,
                                                       EndValue = globalexcel3.EndValue




                                                   });

                    db.SubmitChanges();
                    j = 0;
                    nextrow++;
                }


            }
            this.Progress.Text = "Finished";//update status of insert
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status
        }
        private void PerformanceUpdate(int startrow, int endrow)
        {
            this.Progress.Text = "Working on it...";
            this.Load += new System.EventHandler(this.Form1_Load);

            DataContext db = new DataContext(this.database);

            //create and initialize Excel application, workbook, and worksheet variables
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = null;
            Worksheet sheet = null;

            wkb = excel.Workbooks.Open(exfile,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing); //open Excel file


            sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;//assigns specified worksheet to worksheet variable
            Table<Performance> perf = db.GetTable<Performance>();//creates and initializes LINQ entity table object with table from SQL database



            Range range;

              
            //Delete All Rows
               var perfrows =
               from perfrow in perf
               select perfrow;

            foreach (var perfrow in perfrows)
            {
                perf.DeleteOnSubmit(perfrow);
            }

            try
            {
                db.SubmitChanges();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);         
            }



            for (int j = 1, nextrow = startrow; nextrow < endrow + 1; j++)//starts at startrow and ends at endrow populating global excel object based on corresponding column numbers
            {

                range = sheet.Cells[nextrow, j];




                if (j == 1)
                {
                    globalexcel3.AccountIdentifier = range.Text.ToString();
                }
                else if (j == 2)
                {

                    globalexcel3.FromDate = DateTime.Parse(range.Text.ToString());
                }
                else if (j == 3)
                {
                    globalexcel3.ToDate = DateTime.Parse(range.Text.ToString());
                }
                else if (j == 4)
                {
                    globalexcel3.Currency = range.Text.ToString();
                }
                else if (j == 5)
                {
                    globalexcel3.Category = range.Text.ToString();
                }
                else if (j == 6)
                {
                    globalexcel3.TargetWeight = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 7)
                {
                    globalexcel3.TotalReturn = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 8)
                {
                    globalexcel3.LocalReturn = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 9)
                {
                    globalexcel3.NetReturn = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 10)
                {
                    globalexcel3.BeginMarketValue = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 11)
                {
                    globalexcel3.NetFlowAmount = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 12)
                {
                    globalexcel3.WeightedFlow = System.Convert.ToDecimal(range.Value);
                }
                else if (j == 13)
                {
                    globalexcel3.EndValue = System.Convert.ToDecimal(range.Value);



                    perf.InsertOnSubmit(new Performance
                    {                                   // creates and inserts new Performance object which is one row in the Performance SQL Table
                        AccountIdentifier = globalexcel3.AccountIdentifier,
                        FromDate = globalexcel3.FromDate,
                        ToDate = globalexcel3.ToDate,
                        Currency = globalexcel3.Currency,
                        Category = globalexcel3.Category,
                        TargetWeight = globalexcel3.TargetWeight,
                        TotalReturn = globalexcel3.TotalReturn,
                        LocalReturn = globalexcel3.LocalReturn,
                        NetReturn = globalexcel3.NetReturn,
                        BeginMarketValue = globalexcel3.BeginMarketValue,
                        NetFlowAmount = globalexcel3.NetFlowAmount,
                        WeightedFlowsAmount = globalexcel3.WeightedFlow,
                        EndValue = globalexcel3.EndValue




                    });

                    db.SubmitChanges();
                    j = 0;
                    nextrow++;
                }


            }
        

                this.Progress.Text = "Finished";//update status of update
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status
        }
        private void AssetsInsert(int startrow, int endrow)
        {
            this.Progress.Text = "Working on it...";//update status of insert
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status

            DataContext db = new DataContext(this.database);//creates and initializes DataContext variable using the database connection string

            //create and initialize Excel application, workbook, and worksheet variables
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = null;
            Worksheet sheet = null;

            wkb = excel.Workbooks.Open(exfile,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing);//open Excel file

            sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;//assigns specified worksheet to worksheet variable
            Table<Assets> assets = db.GetTable<Assets>();//creates and initializes LINQ entity table object with table from SQL database
            Range range;
            int nextrow = 0;
 
                nextrow = startrow;
          



            for (int j = 1; nextrow < endrow; j++)//starts at startrow and ends at endrow populating global excel object based on corresponding column numbers
            {
           
                range = sheet.Cells[nextrow, j];




                if (j == 1)
                {
                    globalexcel4.AccountIdentifier = range.Text.ToString();
                }
                else if (j == 2)
                {

                    globalexcel4.TotalMarketValue = System.Convert.ToDecimal(range.Value);
                }
               
                else if (j == 3)
                {
                    globalexcel4.CashBalance = System.Convert.ToDecimal(range.Value);



                    assets.InsertOnSubmit(
                                                   new Assets
                                                   { // creates and inserts new Assets object which is one row in the Assets SQL Table
                                                       AccountIdentifier = globalexcel4.AccountIdentifier,
                                                        TotalMarketValue=globalexcel4.TotalMarketValue,
                                                        CashBalance=globalexcel4.CashBalance




                                                   });

                    db.SubmitChanges();
                    j = 0;
                    nextrow++;
                }


            }
            this.Progress.Text = "Finished";//update status of insert
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status
        }
        private void AssetsUpdate(int startrow, int endrow)
        {
            this.Progress.Text = "Working on it...";//update status of update
            this.Load += new System.EventHandler(this.Form1_Load); //reload form to show progress status

            DataContext db = new DataContext(this.database);//created and initializes DataContext variable using the database connection string

            //create and initialize Excel application, workbook, and worksheet variables
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = null;
            Worksheet sheet = null;

            wkb = excel.Workbooks.Open(exfile,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing);//open Excel file


            sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;//assigns specified worksheet to worksheet variable


            Table<Assets> assets = db.GetTable<Assets>();//creates and initializes LINQ entity table object with table from SQL database
            bool different = false;
            Range range;



            Assets[] E = new Assets[endrow];//creates an array of Assets objects to represent the rows in the Excel file






            for (int i = startrow, j = 1; i < endrow; j++)//starts at startrow and ends at endrow populating global excel object based on corresponding column numbers
            {


                range = sheet.Cells[i, j];



                if (j == 1)
                {
                    globalexcel4.AccountIdentifier = range.Text.ToString();
                }
                else if (j == 2)
                {

                    globalexcel4.TotalMarketValue = System.Convert.ToDecimal(range.Value);
                }

                else if (j == 3)
                {
                    globalexcel4.CashBalance = System.Convert.ToDecimal(range.Value);


                    E[i] = new Assets
                    {// creates a new Assets object which represents a row in the Excel file and stores it in the Excel row object array for later comparision with SQL Table data
                        AccountIdentifier = globalexcel4.AccountIdentifier,
                        TotalMarketValue = globalexcel4.TotalMarketValue,
                        CashBalance = globalexcel4.CashBalance
                    };

                    j = 0;
                    i++;



                }







            }


            int index = startrow;

            foreach (Assets row in assets)
            {


                different = E[index].compare(row);//if the row has been changed in the Excel file then insert the updated Excel row and delete the SQL table row
                if (different)
                {
                    assets.InsertOnSubmit(E[index]);
                    assets.DeleteOnSubmit(row);
                    db.SubmitChanges();
                }
                index++;
            }
            if (index < endrow)
            {

                for (int j = 1; index < endrow; j++)//starts after the last row in the SQL table and ends at endrow populating global excel object based on corresponding column numbers
                {
                    range = sheet.Cells[index, j];






                    if (j == 1)
                    {
                        globalexcel4.AccountIdentifier = range.Text.ToString();
                    }
                    else if (j == 2)
                    {

                        globalexcel4.TotalMarketValue = System.Convert.ToDecimal(range.Value);
                    }

                    else if (j == 3)
                    {
                        globalexcel4.CashBalance = System.Convert.ToDecimal(range.Value);
                    
                        assets.InsertOnSubmit(new Assets
                        {
                            AccountIdentifier = globalexcel4.AccountIdentifier,
                            TotalMarketValue = globalexcel4.TotalMarketValue,
                            CashBalance = globalexcel4.CashBalance
                        });
                        db.SubmitChanges();
                        j = 0;
                        index++;



                    }

                }

               



            }
            this.Progress.Text = "Finished";
            this.Load += new System.EventHandler(this.Form1_Load);
        }
        private void ActivityInsertUpdate(int startrow,int endrow)
        {
            this.Progress.Text = "Working on it...";//update status
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status

            DataContext db = new DataContext(this.database);//created and initializes DataContext variable using the database connection string

            //create and initialize Excel application, workbook, and worksheet variables
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = null;
            Worksheet sheet = null;

            wkb = excel.Workbooks.Open(exfile,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing);//open Excel file

            sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;//assigns specified worksheet to worksheet variable
            Table<Activity> activity = db.GetTable<Activity>(); //creates and initializes LINQ entity table object with table from SQL database



            Range range;



            int nextrow = 0;


            //in order to update, the row in the Activity SQL Table with the ActivityID of the greatest value is queried and its ActivityID is used to locate the location in the Excel file to begin updating from
            if (Update.Checked)
            {
                  IEnumerable<Activity> query = activity;

                  int result = query.Max<Activity,int>(row => row.ActivityID);
                bool done = false;
                int excelactivityid = 0; 
            for(int k=startrow;done==false;k++)
                {
                    range = sheet.Cells[k, 2];
                    excelactivityid=Int32.Parse(range.Text.ToString());
                    if(excelactivityid==result)
                    {
                        nextrow = k+1;
                        done = true;
                        endrow = endrow + 1;
                    }
                }

            }
            else
            {

                nextrow = startrow;
            }



            //starts after the row with the ActivityID of the greatest value or at startrow's value and ends at endrow populating 
            //global excel object based on corresponding column numbers
            for (int j = 1; nextrow < endrow; j++)
            {

                range = sheet.Cells[nextrow, j];


              
                if (j == 1)
                {
                    globalexcel5.BusinessID = Int32.Parse(range.Text.ToString());
                }
                else if (j == 2)
                {

                    globalexcel5.ActivityID = Int32.Parse(range.Text.ToString());
                }
                else if(j==3)
                {
                    globalexcel5.Business = range.Text.ToString();
                }
                else if (j == 4)
                {
                    globalexcel5.Rep = range.Text.ToString();
                }
                else if (j == 5)
                {
                    globalexcel5.Rep2 = range.Text.ToString();
                }
                else if (j == 6)
                {
                    globalexcel5.InvTeam1 = range.Text.ToString();
                }
                else if (j == 7)
                {
                    globalexcel5.Invteam2 = range.Text.ToString();
                }
                else if (j == 8)
                {
                    globalexcel5.StartDate = DateTime.Parse(range.Text.ToString());
                }
                else if (j == 9)
                {
                    globalexcel5.ActivityType = range.Text.ToString();



                    activity.InsertOnSubmit(
                                                   new Activity
                                                   {
                                                       BusinessID = globalexcel5.BusinessID,
                                                        ActivityID = globalexcel5.ActivityID,
                                                        Business=globalexcel5.Business,
                                                        Rep=globalexcel5.Rep,
                                                        Rep2=globalexcel5.Rep2,
                                                        InvTeam1=globalexcel5.InvTeam1,
                                                        InvTeam2=globalexcel5.Invteam2,
                                                        StartDate=globalexcel5.StartDate,
                                                        ActivityType=globalexcel5.ActivityType




                                                   });

                    db.SubmitChanges();
                    j = 0;
                    nextrow++;
                }


            }
            this.Progress.Text = "Finished";//update status
            this.Load += new System.EventHandler(this.Form1_Load);//reload form to show progress status
        }

       
    }
    }
    
    



