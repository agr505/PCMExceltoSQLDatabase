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
    public class globalexcel5
    {
        public static int BusinessID,ActivityID;
        public static string Business, Rep, Rep2, InvTeam1, Invteam2, ActivityType;
        public static DateTime StartDate;

    }
    [Table(Name = "Activity")]//Name of SQL Table this Object should be mapped to
    class Activity
    {
        //Object properties that will be mapped to SQL Table Columns
        private int _BusinessID;
        [Column( Storage = "_BusinessID")]
        public int BusinessID
        {
            get { return this._BusinessID; }
            set { this._BusinessID = value; }
        }

        private int _ActivityID;
        [Column(IsPrimaryKey = true, Storage = "_ActivityID")]
        public int ActivityID
        {
            get { return this._ActivityID; }
            set { this._ActivityID = value; }
        }
        private string _Business;
        [Column( Storage = "_Business")]
        public string Business
        {
            get { return this._Business; }
            set { this._Business = value; }
        }
    
    private string _Rep;
    [Column(Storage = "_Rep")]
    public string Rep
    {
        get { return this._Rep; }
        set { this._Rep = value; }
    }
        private string _Rep2;
        [Column(Storage = "_Rep2")]
        public string Rep2
        {
            get { return this._Rep2; }
            set { this._Rep2 = value; }
        }

        private string _InvTeam1;
[Column( Storage = "_InvTeam1")]
public string InvTeam1
        {
        get { return this._InvTeam1; }
set { this._InvTeam1 = value; }
    } 

private string _InvTeam2;
[Column( Storage = "_InvTeam2")]
public string InvTeam2
        {
        get { return this._InvTeam2; }
set { this._InvTeam2 = value; }
    }

        private DateTime _StartDate;
        [Column(Storage = "_StartDate")]
        public DateTime StartDate
        {
            get { return this._StartDate; }
            set { this._StartDate = value; }
        }
        private string _ActivityType;
        [Column(Storage = "_ActivityType")]
        public string ActivityType
        {
            get { return this._ActivityType; }
            set { this._ActivityType = value; }
        }

    }
}
