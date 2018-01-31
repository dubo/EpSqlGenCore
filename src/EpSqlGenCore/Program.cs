using System;
using EpSqlGen;
using EpSqlGenCore;

namespace EpSqlGen
{
    class Program
    {
        static void Main(string[] args)
        {             
            var myReport = new EpSqlGenerator(args);
            // Logging only after initialization
            SimpleLog.WriteLog("");
            SimpleLog.WriteLog("***  Excel generator started  ***");
            myReport.SetupEnviroment();

            using (var conn = DbTools.CreateOpenedDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings[myReport.repConnectString].ProviderName, System.Configuration.ConfigurationManager.ConnectionStrings[myReport.repConnectString].ConnectionString))
            {

                if (myReport.outFormat == ".json")
                    myReport.GenerateJson(conn);
                else
                    myReport.GenerateXlsxReport(conn);
            }
            myReport.CloseEnviroment();
           
        }
    }
}
