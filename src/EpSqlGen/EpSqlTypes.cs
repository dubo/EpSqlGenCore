using System;
using System.Collections.Generic;

namespace EpSqlGen
{
    public enum ExitCodes : int
    {
        Success = 0,
        DbConnectionFailed = 1,
        MissingParams = 2,
        MissingWorkingDirs = 3,
        MissingFiles = 4,
        MissingTemplate = 5,
        MissingConfiguration = 6,
        UnknownError = 999
    }

    public class Argument
    {
        public string name { get; set; }
        public string type { get; set; }
        public string value { get; set; }
    }

    public class TableFields
    {
        public int colId { get; set; }
        public string title { get; set; }
        public string name { get; set; }
        public string type { get; set; }
        public string format { get; set; }
        public int minsize { get; set; }
        public int order { get; set; }
    }

    public class Tab
    {
        public string query { get; set; }
        public string name { get; set; }
        public string title { get; set; }
        public Boolean printHeader { get; set; }
        public List<TableFields> fields { get; set; }
    }

    public class ExcelDef
    {
        public string fileName { get; set; }
        public string template { get; set; }
        public Boolean autofilter { get; set; }
        public Boolean timestamp { get; set; }
        public string version { get; set; }
        public List<Tab> tabs { get; set; }
    }

}

