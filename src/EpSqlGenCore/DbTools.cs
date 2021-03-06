﻿using System;
using EpSqlGen;
using System.Data.Common;
using System.Data.OracleClient;
//using System.Data.SQLite;
//using SQLitePCLRaw.bundle_e_sqlite3
//using Microsoft.Data.Sqlite.Core  - this time same required methods is missing


namespace EpSqlGenCore
{
    public class DbTools
    {

        //DB setup
        //https://stackoverflow.com/questions/9218847/how-do-i-handle-database-connections-with-dapper-in-net
        //https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/obtaining-a-dbproviderfactory
        // Given a provider name and connection string, 
        // create the DbProviderFactory and DbConnection.
        // Returns a DbConnection on success; null on failure.
        public static System.Data.Common.DbConnection CreateOpenedDbConnection(string providerName, string connectionString)
        {
            // Assume failure.
            DbConnection connection = null;

            // Create the DbProviderFactory and DbConnection.
            if (connectionString != null)
            {
                DbProviderFactory factory;
                try
                {
                    if (providerName == "Mono.Data.OracleClientCore" || providerName == "System.Data.OracleClient")
                        factory = OracleClientFactory.Instance;
                    else if (providerName == "Oracle.ManagedDataAccess.Client")
                        factory = Oracle.ManagedDataAccess.Client.OracleClientFactory.Instance ;
                    else if (providerName == "Npgsql")
                        factory = Npgsql.NpgsqlFactory.Instance;
                    else if (providerName == "System.Data.SqlClient")
                        factory = System.Data.SqlClient.SqlClientFactory.Instance;
                    //    Microsoft.Data.Sqlite  disabled, becose in .Net Core method GetSchemaTable() not implemented
                     else if (providerName == "Microsoft.Data.Sqlite")
                     {
                         factory = Microsoft.Data.Sqlite.SqliteFactory.Instance;
                         SQLitePCL.raw.SetProvider(new SQLitePCL.SQLite3Provider_e_sqlite3());
                     }     
                    else if (providerName == "MySql.Data")
                        factory = MySql.Data.MySqlClient.MySqlClientFactory.Instance;
                    else
                        factory = null;

                    if (factory != null)
                    {
                        connection = factory.CreateConnection();
                        connection.ConnectionString = connectionString;
                        connection.Open();
                    }
                    else
                        throw new Exception("Unknown DB Provider: " + providerName);
                }
                catch (Exception ex)
                {
                    // Set the connection to null if it was created.
                    if (connection != null)
                    {
                        connection = null;
                    }
                    Console.WriteLine(ex.Message);
                    SimpleLog.WriteLog("****  Openinig DB connection with connection failed. Provider: " + providerName);
                    SimpleLog.WriteLog(ex.Message);
                    System.Environment.Exit((int)ExitCodes.DbConnectionFailed);
                }
            }
            // Return the connection.
            return connection;
        }

    }
}
