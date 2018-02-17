using Dapper;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace EpSqlGen
{

    public class EpSqlGenerator
    {
        private string outFileName = null;
        private string defFileName = null;
        private DirectoryInfo inputDir;
        private DirectoryInfo outputDir;
        private FileInfo newFile;
        private FileInfo templateFile = null;
        private ExcelDef def;
        private ExcelDef outDef;
        private Dictionary<String, Object> arguments = new Dictionary<String, Object>();         
        private char dbParamChar = '@';    
        private bool enabledTimestamp = true;
        private bool enabledConsoleLog = false;
        private bool enabledOutput = false;
        private bool consoleMode = false;
        public string outFormat = ".xlsx";
        public string repConnectString;

        public static void WriteHelp()
        {            
            var appName = System.AppDomain.CurrentDomain.FriendlyName;
            Console.WriteLine(".Net (.Net core) xlsx/json generator for Sql (now only Oracle tested)");
            Console.WriteLine("");
            Console.WriteLine("Switches:");
            Console.WriteLine("-h                   // Display help info");
            Console.WriteLine("-v                   // Version ");
            Console.WriteLine("-j                   // Generate json output file, (xlsx file is default)");
            Console.WriteLine("-jc                  // Generate json output only to console, usefull for integration with other products ");
            Console.WriteLine("-ec                  // Enable console logging output - usefull for debuggnig");
            Console.WriteLine("-eo                  // Enable  out file generation -  created definition  in json format ");
            Console.WriteLine("-dt                  // Disable timestamp mark in generated xlsx file");
            Console.WriteLine("-c:MyConnectionStrig // Connect string name as argument");
            Console.WriteLine("-a:MyArgumentName:MyArgumentType:MyargumentValue     // Argumet for SQL    ");
            Console.WriteLine("");
            Console.WriteLine("Usage(sql definition file for simple one tab output):");
            Console.WriteLine("dotnet " + appName + ".dll MySqlQuery.sql -oMyOutputFileName -a:MyArgument1:Argument1Type:Argumet1value -a:MyArgument2:Argument2Type:Argumet2value");
            Console.WriteLine("");
            Console.WriteLine("Usage(json definition file for complex output):");
            Console.WriteLine("dotnet " + appName + ".dll MyJsonDefinition.json -oMyOutputFileName -aMyArgument1:Argument1Type:Argumet1value -aMyArgument2:Argument2Type:Argumet2value");
            Console.WriteLine("");
            Console.WriteLine("Supported arguments types:  string || char || varchar2 || varchar || date || integer || decimal || number || array ");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Sample-sql src to  XLSX (portable/Win exe sample): ");
            Console.WriteLine("dotnet " + appName + ".dll MySqlQuerry.sql -oMyOutputfile  -dt -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Zmluva:number:505115 ");
            Console.WriteLine( appName + ".exe MySqlQuerry.sql -oMyOutputfile  -dt -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Zmluva:number:505115 ");
            Console.WriteLine("");
            Console.WriteLine("Sample-json src to XLSX (portable/Win exe sample):") ;
            Console.WriteLine("dotnet " + appName + ".dll Test.json -oMyOutputfile -do -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Zmluva:number:505115 ");
            Console.WriteLine( appName + ".exe Test.json -oMyOutputfile -do -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Zmluva:number:505115 ");
            Console.WriteLine("");
            Console.WriteLine("Sample-sql to json  (portable/Win exe sample):");
            Console.WriteLine("dotnet " + appName + ".dll MySqlQuerry.sql -oMyOutputfile -j -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Contract:integer:505115");
            Console.WriteLine( appName + ".exe MySqlQuerry.sql -oMyOutputfile -j -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Contract:integer:505115");
        }

        public EpSqlGenerator(string[] inputArgs)
        {
            //:arStavy:array:'P9','K9','O9' :asProdukt:string:UO :adOd:date:4.2.2015 :anZmluva:number:505115
            // Test.sql -oMyOutputfile  -aarStavy:array:'P9','K9','O9' -aanStavy:array:31,32,33 -aasProdukt:string:UO -aadOd:date:4.2.2015 -aanZmluva:number:505115
            foreach (var arg in inputArgs)
            {
                if (arg[0] == '-')
                {
                    if (arg[1] == 'o')
                        outFileName = arg.Substring(2).Trim();
                    else if (arg[1] == 'c')
                        repConnectString = (arg[2] == ':') ? arg.Substring(3).Trim() : arg.Substring(2).Trim();
                    else if (arg[1] == 'j')
                    {
                        outFormat = ".json";
                        if (arg.Length == 3 && arg[2] == 'c')
                        {
                            SimpleLog.DisableConsole(true);
                            consoleMode = true;
                            enabledOutput = false;
                        }
                    }
                    else if (arg[1] == 'e' && arg.Length == 3 && arg[2] == 'o')
                        enabledOutput = true;
                    else if (arg[1] == 'e' && arg.Length == 3 && arg[2] == 'c')
                    {
                        SimpleLog.DisableConsole(true);
                        enabledConsoleLog = true;
                    }
                    else if (arg[1] == 'd' && arg.Length == 3 && arg[2] == 't')
                        enabledTimestamp = false;
                    else if (arg[1] == 'h')
                    {
                        WriteHelp();
                        System.Environment.Exit(0);
                    }
                    else if (arg[1] == 'v')
                    {
                        Console.WriteLine(System.AppDomain.CurrentDomain.FriendlyName  + " version: " + System.Reflection.Assembly.GetEntryAssembly().GetName().Version )  ;
                        string osplatform;
                        if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows)) 
                            osplatform = "Windows";
                        else if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX)) 
                            osplatform = "OSX";
                        else if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Linux)) 
                            osplatform = "Linux";
                        else
                            osplatform = "Unknown";
                        Console.WriteLine("Enviroment OS: " + Environment.OSVersion.Platform + ", platform ID " + (int)Environment.OSVersion.Platform );
                        Console.WriteLine("OsPlatform   : " + osplatform);
                        Console.WriteLine("Architecture : " + System.Runtime.InteropServices.RuntimeInformation.OSArchitecture);
                        Console.WriteLine("OsDescription: " + System.Runtime.InteropServices.RuntimeInformation.OSDescription);
                        System.Environment.Exit(0);
                    }                       
                    else if (arg[1] == 'a')
                    {
                        string[] argdetails = (arg[2] == ':') ? arg.Substring(3).Split(':') : arg.Substring(2).Split(':');
                        if (argdetails[1].ToLower() == "string" || argdetails[1].ToLower().Contains("char"))
                            arguments.Add(argdetails[0], argdetails[2]);
                        else if (argdetails[1].ToLower().Contains("date"))
                            arguments.Add(argdetails[0], DateTime.Parse(argdetails[2]));
                        else if (argdetails[1].ToLower() == "integer" || argdetails[1].ToLower() == "int")
                            arguments.Add(argdetails[0], int.Parse(argdetails[2]));
                        else if (argdetails[1].ToLower() == "decimal" || argdetails[1].ToLower() == "number")
                            arguments.Add(argdetails[0], decimal.Parse(argdetails[2]));
                        else if (argdetails[1].ToLower().Contains("array"))
                        {
                            var items = argdetails[2].Split(',');
                            if (argdetails[2].Contains("'") || argdetails[1].ToLower().Contains("string") || argdetails[1].ToLower().Contains("char"))
                            {
                                List<string> vals = new List<string>();
                                foreach (var item in items)
                                    vals.Add(item.Replace("'", ""));
                                arguments.Add(argdetails[0], vals.ToArray());
                            }
                            else
                            {
                                List<int> vals = new List<int>();
                                foreach (var item in items)
                                    vals.Add(int.Parse(item));
                                arguments.Add(argdetails[0], vals.ToArray());
                            }
                        }
                    }
                }
                else
                    // ide o naazov suboru s definiciou
                    if (arg.ToLower().IndexOf(".sql") > 0 || arg.ToLower().IndexOf(".json") > 0)
                    defFileName = arg;
            }

            // Logging only after initialeze console output
            SimpleLog.WriteLog("Input arguments {");
            foreach (var arg in inputArgs)
                SimpleLog.WriteLog(arg, false);
            SimpleLog.WriteLog("}", false);
            SimpleLog.WriteLog("", false);
            if (repConnectString == null)
                repConnectString = "ReportConnString";
        }

        private void SetDbParams(System.Data.Common.DbConnection conn)
        {
            if (conn.GetType().Name.ToLower().IndexOf("oracle") > -1 || conn.GetType().Name.ToLower().IndexOf("pgsql") > -1)            
                dbParamChar = ':';               
        }

        private Dictionary<String, Object> QuerryArguments(string querry)
        {
            var qArgs = new Dictionary<String, Object>();
            var normQ = querry.Replace("\\r\\n", " ").Replace("\r\n", " ").Replace(")", " ").Replace(",", " ").Replace("|", " ");
            foreach (var arg in arguments)
                if (normQ.IndexOf(dbParamChar + arg.Key + ' ') > -1)
                    qArgs.Add(arg.Key, arg.Value);

            return qArgs;
        }

        public void SetupEnviroment()
        {
            // Check definition file
            if (defFileName == null)
            {
                SimpleLog.WriteException("****  EpSqlGen aborted, missing definition file in arguments ");
                System.Environment.Exit((int)ExitCodes.MissingParams);
            }
            else
            {
                // definition file in actual directory ?
                FileInfo definitionFile = new FileInfo(defFileName);
                if (!definitionFile.Exists && Path.GetDirectoryName(defFileName) == "")
                {
                    // Check work directiries
                    var defDir = System.Configuration.ConfigurationManager.AppSettings.Get("DefinitionsDir");
                    if (defDir == null || defDir == "")
                    {
                        SimpleLog.WriteException("****  EpSqlGen aborted, definition for Definitions directory does not exist!");
                        System.Environment.Exit((int)ExitCodes.MissingConfiguration);
                    }
                    else
                        inputDir = new DirectoryInfo(defDir);

                    if (!inputDir.Exists)
                    {
                        SimpleLog.WriteException("****  EpSqlGen aborted, definitions for input directory does not exist!");
                        System.Environment.Exit((int)ExitCodes.MissingWorkingDirs);
                    }
                    // look in input dir from setup
                    definitionFile = new FileInfo(Path.Combine(inputDir.FullName, defFileName));
                }
                else
                    inputDir = new DirectoryInfo(Path.GetDirectoryName(definitionFile.FullName));


                if (!definitionFile.Exists)
                {
                    SimpleLog.WriteException("****  EpSqlGen aborted, definitions file does not exist!");
                    System.Environment.Exit((int)ExitCodes.MissingFiles);
                }
                else
                {
                    // Open def
                    string defFileContent = File.ReadAllText(definitionFile.FullName);
                    if (definitionFile.Extension == ".sql")
                    {
                        def = new ExcelDef { fileName = definitionFile.Name.Substring(0, definitionFile.Name.ToLower().IndexOf(".sql")), timestamp = true, autofilter = true, tabs = new List<Tab> { } };
                        def.tabs.Add(new Tab { name = (outFormat == ".json" ? "rows" : def.fileName), title = "", query = defFileContent.TrimEnd() });
                    }
                    else
                        def = JsonConvert.DeserializeObject<ExcelDef>(defFileContent);

                }

                outDef = new ExcelDef { fileName = def.fileName, template = def.template, autofilter = def.autofilter, timestamp = def.timestamp, tabs = new List<Tab>() };

                // Verify template file existence
                if (!(def.template == null || def.template == ""))
                {
                    templateFile = new FileInfo(inputDir.FullName);
                    if (!templateFile.Exists)
                    {
                        if (!inputDir.Exists)
                        {
                            SimpleLog.WriteException("****  EpSqlGen aborted, definitions for input directory does not exist!");
                            System.Environment.Exit((int)ExitCodes.MissingWorkingDirs);
                        }
                        templateFile = new FileInfo(Path.Combine(inputDir.FullName, def.template));
                    }

                    if (!templateFile.Exists)
                    {
                        SimpleLog.WriteException("Template file: " + def.template + " does not exist! (" + templateFile.FullName + ")");
                        System.Environment.Exit((int)ExitCodes.MissingFiles);
                    }
                }

                // Check how output file is defined - wih full path or file name only 
                if (outFileName == "" || outFileName == null)
                {
                    // if outfile in params is empty , check json definition  ->  def.fileName
                    if (definitionFile.Extension == ".sql")
                        outFileName = Path.GetFileNameWithoutExtension(defFileName);
                    else
                        outFileName = ((def.fileName == "" || def.fileName == null) ? Path.GetFileNameWithoutExtension(defFileName) : def.fileName);

                    if (new FileInfo(defFileName).Exists)
                        outputDir = new DirectoryInfo(Path.GetDirectoryName(definitionFile.FullName));
                }

                if (outputDir == null)
                {
                    if (Path.GetDirectoryName(outFileName) != "")
                        outputDir = new DirectoryInfo(Path.GetDirectoryName(outFileName));
                    else
                    {
                        var outDir = System.Configuration.ConfigurationManager.AppSettings.Get("OutputsDir");
                        if (outDir == null || outDir == "")
                        {
                            SimpleLog.WriteException("****  EpSqlGen aborted, definition for Definitions directory does not exist!");
                            System.Environment.Exit((int)ExitCodes.MissingConfiguration);
                        }
                        outputDir = new DirectoryInfo(outDir);
                    }
                }

                // Verify output file existence
                if (!outputDir.Exists)
                {
                    SimpleLog.WriteException("****  EpSqlGen aborted, definitions for output directory does not exist! (" + outputDir.FullName + ")");
                    SimpleLog.WriteException(outputDir.FullName);
                    System.Environment.Exit((int)ExitCodes.MissingWorkingDirs);
                }

                // switch mode, if output file is defined as json
                if (Path.GetExtension(outFileName) == ".json" && outFormat != ".json")
                    outFormat = ".json";

                // fix extension
                if (Path.GetExtension(outFileName) != outFormat)
                    outFileName = outFileName + outFormat;

                // Create output file  
                newFile = new FileInfo(Path.Combine(outputDir.FullName, Path.GetFileName(outFileName)));
                if (newFile.Exists)
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(Path.Combine(outputDir.FullName, Path.GetFileName(outFileName)));
                }
                SimpleLog.WriteLog("Started generate output for " + defFileName + " definition file.");
            }
        }

        public void CloseEnviroment()
        {
            if (enabledOutput && outDef.tabs != null)
            {

                FileInfo outConfName = new FileInfo(Path.Combine(outputDir.FullName, Path.GetFileNameWithoutExtension(defFileName) + ".out"));
                if (outConfName.Exists)
                {
                    outConfName.Delete();
                    outConfName = new FileInfo(Path.Combine(outputDir.FullName, Path.GetFileNameWithoutExtension(defFileName) + ".out"));
                }

                using (StreamWriter sw = outConfName.CreateText())
                {
                    // Replace("\\r\\n", "\r\n") - need for readable SQL in JSON
                    sw.Write(JsonConvert.SerializeObject(outDef, Formatting.Indented).Replace("\\r\\n", "\r\n"));
                    // backup for PrettyPrintJson
                    //sw.Write(JsonSerializer.SerializeToString(outDef).PrettyPrintJson().Replace("\\r\\n", "\r\n"));
                }
                SimpleLog.WriteLog(" Used configuration saved into file " + outConfName.ToString());
            }

            SimpleLog.WriteLog("***  End generating output for " + defFileName + " definition file ****");
            SimpleLog.WriteLog("", false);
        }

        private ExcelWorksheet GetWs(ExcelPackage package, Tab tab)
        {
            ExcelWorksheet ws;

            if (templateFile == null)
                ws = package.Workbook.Worksheets.Add(tab.name);
            else
            {
                ws = package.Workbook.Worksheets[tab.name];
                if (ws == null)
                {
                    SimpleLog.WriteException("Worksheet " + tab.name + " does not exists in template" + def.template);
                    System.Environment.Exit((int)ExitCodes.MissingTemplate);
                }

            }
            return ws;
        }


        /* disabled after switch from ServiceSstack to Dapper only 
        // http://www.codeproject.com/Articles/80343/Accessing-private-members.aspx#_comments
        // http://stackoverflow.com/questions/3303126/how-to-get-the-value-of-private-field-in-c
        // + neskor pomocka http://stackoverflow.com/questions/6563470/can-a-c-sharp-method-return-a-method
        internal static object GetInstanceField(Type type, object instance, string fieldName)
        {
            BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
                | BindingFlags.Static;
            FieldInfo field = type.GetField(fieldName, bindFlags);
            return field.GetValue(instance);
        }

        */

        private static string getVariableType(String fullName, String decimalSize)
        {
            String origType;
            if (fullName == null)
                origType = "Uknown";
            // disabled after switch from SSstack to Dapper only
            // else if (bb.GetType().ToString() == "ServiceStack.OrmLite.Oracle.OracleValue")
            // type = GetInstanceField(typeof(ServiceStack.OrmLite.Oracle.OracleValue), bb, "_oracleValueType").ToString();
            if (fullName.LastIndexOf('.') > 0)
                origType = fullName.Substring(fullName.LastIndexOf('.') + 1);
            else
                origType = fullName;

            if (origType == "Decimal")
                return decimalSize == "0" ? "Integer" : origType;
            return origType;
        }

        private static List<TableFields> getDeclaration(System.Data.IDataReader reader, List<TableFields> fields)
        {
            var table = reader.GetSchemaTable();
            var actualFields = new List<TableFields> { };
            var orderedFields = new List<TableFields> { };
            TableFields item;
            String dateFormat;

            switch (System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)
            {
                case "d.M.yyyy":
                case "d. M. yyyy":
                    dateFormat = "d.m.yyyy";
                    break;
                default:
                    dateFormat = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
                    break;
            }

            for (int i = 0; i < reader.FieldCount; i++)
            {
                System.Data.DataRow row = table.Rows[i];
                actualFields.Add(new TableFields { colId = i, name = row["ColumnName"].ToString(), type = getVariableType(row["DataType"].ToString(), row["NumericScale"].ToString()) });

                // ak by som nemal ziadnu definiciu predtym
                item = (fields == null ? null : fields.FirstOrDefault(x => x.name.ToUpper() == row["ColumnName"].ToString().ToUpper() ) );
                if (item == null)
                {
                    actualFields[i].title = actualFields[i].name;
                    actualFields[i].minsize = Math.Max(Math.Min(10, int.Parse(row["ColumnSize"].ToString())), actualFields[i].name.Length);
                    actualFields[i].order = 100 + i;
                    actualFields[i].format = actualFields[i].type == "DateTime" ? dateFormat : "auto";
                }
                else
                {
                    actualFields[i].title = (item.title == null ? actualFields[i].name : item.title).Trim();
                    actualFields[i].format = item.format == null ? (actualFields[i].type == "DateTime" ? dateFormat : "auto") : item.format;
                    actualFields[i].minsize = item.minsize == 0 ? Math.Max(Math.Min(10, int.Parse(row["ColumnSize"].ToString())), actualFields[i].name.Length) : item.minsize;
                    actualFields[i].order = item.order;
                }
            }

            int j = 0;
            int k = 0;
            foreach (var field in actualFields.OrderBy(o => o.order).ToList())
            {
                orderedFields.Add(field);
                if (field.order != 0)
                {
                    j++;
                    orderedFields[k].order = j;
                }
                k++;
            }

            SimpleLog.WriteLog(JsonConvert.SerializeObject(orderedFields, Formatting.Indented));
            return orderedFields;
        }

        // Json generator https://stackoverflow.com/questions/5083709/convert-from-sqldatareader-to-json
        private Dictionary<string, object> SerializeRow(List<TableFields> cols, System.Data.IDataReader reader)
        {
            var result = new Dictionary<string, object>();
            foreach (var col in cols)
                if (col.order != 0)
                    result.Add(col.title, col.type == "Integer" ? reader.GetInt32(col.colId) : reader[col.colId]);
            return result;
        }

        public void GenerateJson(System.Data.Common.DbConnection conn)
        {
            var output = new Dictionary<String, Object>();
            foreach (var tab in def.tabs)
            {
                System.Data.IDataReader reader = conn.ExecuteReader(tab.query, QuerryArguments(tab.query));
                var results = new List<Dictionary<string, object>>();
                List<TableFields> rowConfig = getDeclaration(reader, tab.fields);
                outDef.tabs.Add(new Tab { name = tab.name, title = tab.title, query = tab.query, fields = rowConfig });
                while (reader.Read())
                    results.Add(SerializeRow(rowConfig, reader));
                output.Add(tab.name, results);
            }

            // Create output file or print to console
            if (!consoleMode)
            {
                newFile = new FileInfo(Path.Combine(outputDir.FullName, Path.GetFileName(outFileName)));
                if (newFile.Exists)
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(Path.Combine(outputDir.FullName, Path.GetFileName(outFileName)));
                }
                using (StreamWriter sw = newFile.CreateText())
                {
                    sw.Write(JsonConvert.SerializeObject(output, Formatting.Indented));
                }
                // LOg or return out file 
                if (!enabledOutput)
                    Console.Write(JsonConvert.SerializeObject(new { output_file = newFile.FullName }, Formatting.Indented));
                else
                    SimpleLog.WriteLog("Generated file: " + newFile.FullName);
            }
            else
                Console.Write(JsonConvert.SerializeObject(output, Formatting.Indented));
        }

        public void GenerateXlsxReport(System.Data.Common.DbConnection conn)
        {
            SetDbParams(conn);
            using (ExcelPackage package = (templateFile != null ? new ExcelPackage(newFile, templateFile) : new ExcelPackage(newFile)))
            {
                foreach (var tab in def.tabs)
                {
                    var ws = GetWs(package, tab);
                    var riadok = 1;
                    var stlpec = 1;
                    var start = 1;


                    // set start  position for Label - that would be Global or Local
                    if (templateFile != null)
                    {
                        ExcelNamedRange label = null;
                        try
                        {
                            label = ws.Names["Label"];
                        }
                        catch
                        {
                            //Console.WriteLine("{0} Exception caught.", e);
                            try
                            {
                                label = package.Workbook.Names["Label"];
                            }

                            catch
                            {
                                SimpleLog.WriteLog("Label field not found in this workbook/template");
                            }
                        }

                        if (label != null)
                        {
                            riadok = label.Start.Row;
                            stlpec = label.Start.Column;
                        }
                    }

                    var nadpis = "";
                    if (!(tab.title == null || tab.title == ""))
                    {                       
                        if (tab.title.Trim().ToUpper().Substring(0, 6) == "SELECT")
                            nadpis = conn.QuerySingle<string>(tab.title, QuerryArguments(tab.title));
                        else
                            nadpis = tab.title;
                    }

                    // Main Select
                    // https://github.com/ericmend/oracleClientCore-2.0/blob/master/test/dotNetCore.Data.OracleClient.test/OracleClientCore.cs
                    System.Data.IDataReader reader = conn.ExecuteReader(tab.query, QuerryArguments(tab.query));
          
                    List <TableFields> rowConfig = getDeclaration(reader, tab.fields);
                    outDef.tabs.Add(new Tab { name = tab.name, title = tab.title, query = tab.query, fields = rowConfig });

                    int r = 0;
                    int activeColCount = 0;
                    while (reader.Read())
                    {
                        r++;
                        //Initial section for sheet  
                        if (r == 1)
                        {
                            if (nadpis != null && nadpis != "")
                            {
                                ws.Cells[riadok, stlpec].Value = nadpis;

                                if (templateFile == null)
                                {
                                    using (ExcelRange rr = ws.Cells[riadok, stlpec, riadok, stlpec - 1 + rowConfig.Where(o => o.order != 0).Count()])
                                    {
                                        rr.Merge = true;
                                        //rr.Style.Font.SetFromFont(new Font("Britannic Bold", 12, FontStyle.Italic));
                                        rr.Style.Font.Size = 12;
                                        rr.Style.Font.Bold = true;
                                        rr.Style.Font.Color.SetColor(Color.FromArgb(63, 63, 63));
                                        rr.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                                        rr.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        //r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                                        rr.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
                                    }
                                }
                                riadok++;
                            }

                            // set Data position
                            if (templateFile != null)
                            {
                                ExcelNamedRange label = null;
                                try
                                {
                                    label = ws.Names["Data"];
                                }
                                catch
                                {
                                    try
                                    {
                                        label = package.Workbook.Names["Data"];
                                    }
                                    catch
                                    {
                                        SimpleLog.WriteLog("Data field not found in this workbook/template");
                                    }
                                }

                                if (label != null)
                                {
                                    riadok = label.Start.Row - 1; // Header je nad riadkom  (above row)
                                    stlpec = label.Start.Column;
                                }
                            }

                            //Add the headers                                 
                            for (int i = 0; i < rowConfig.Count; i++)
                                if (rowConfig[i].order != 0)
                                {
                                    activeColCount++;
                                    if (templateFile == null || tab.printHeader)
                                        ws.Cells[riadok, activeColCount + stlpec - 1].Value = rowConfig[i].title;
                                    ws.Names.Add(rowConfig[i].name, ws.Cells[riadok, activeColCount + stlpec - 1]);
                                }

                            //Ok now format the values;
                            //ws.Cells[1, 1, 1, tab.fields.Count].Style.Font.Bold = true; //Font should be bold
                            //ws.Cells[1, 1, 1, tab.fields.Count].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            //ws.Cells[1, 1, 1, tab.fields.Count].Style.Fill.BackgroundColor.SetColor(Color.Aqua);

                            if (templateFile == null || tab.printHeader)
                            {
                                using (var range = ws.Cells[riadok, stlpec, riadok, activeColCount + stlpec - 1])
                                {
                                    range.Style.Font.Color.SetColor(Color.White);
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));
                                    range.Style.Font.Bold = true;
                                    range.Style.WrapText = true;
                                    //Only need to grab the first cell of the merged range
                                    //ws.Cells[$"A{row}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    //ws.Cells[$"A{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                }
                            }
                            riadok++;
                            // set line for data 
                            start = riadok;
                        }

                        //data section, FETCH RECORDS                                         
                        for (int i = 0; i < rowConfig.Count(); i++)
                        {
                            var colId = rowConfig[i].colId;
                            if (rowConfig[i].order != 0 && !reader.IsDBNull(colId))
                            {
                                var a = reader.GetValue(colId);
                                switch (rowConfig[i].type)
                                {
                                    case "String":
                                        var pom = reader.GetString(colId);
                                        ws.SetValue(riadok, rowConfig[i].order + stlpec - 1, pom.Substring(0, pom.Length  /*- correctStringChars */ ));
                                        break;
                                    case "Integer":
                                        ws.SetValue(riadok, rowConfig[i].order + stlpec - 1, reader.GetInt32(colId));
                                        break;
                                    case "DateTime":
                                        ws.SetValue(riadok, rowConfig[i].order + stlpec - 1, reader.GetDateTime(colId));
                                        break;
                                    case "Decimal":
                                        ws.SetValue(riadok, rowConfig[i].order + stlpec - 1, reader.GetDecimal(colId));
                                        break;
                                    case "Byte[]":
                                        ws.SetValue(riadok, rowConfig[i].order + stlpec - 1, reader.GetByte(colId));
                                        break;
                                    default:
                                        ws.SetValue(riadok, rowConfig[i].order + stlpec - 1, reader.GetValue(colId).ToString());
                                        break;
                                };
                            }
                        }
                        riadok++;

                    }

                    // no rows
                    if (r == 0)
                    {
                        if (nadpis != null)
                        {
                            ws.Cells[riadok, stlpec].Value = nadpis;
                            using (ExcelRange rr = ws.Cells[riadok, stlpec, riadok, 8 + stlpec])
                            {
                                rr.Merge = true;
                                rr.Style.Font.Size = 12;
                                rr.Style.Font.Bold = true;
                                rr.Style.Font.Color.SetColor(Color.FromArgb(63, 63, 63));
                                rr.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                                rr.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                rr.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
                            }
                            riadok++;
                        }
                        ws.Cells[riadok, stlpec].Value = "No rows in querry result";
                        SimpleLog.WriteLog("No rows in Query from  " + tab.name);
                    }
                    else
                    {
                        foreach (var row in rowConfig.Where(o => o.format != null && o.format != "auto"))
                            ws.Cells[start, row.order + stlpec - 1, riadok, row.order + stlpec - 1].Style.Numberformat.Format = row.format;
                        //ws.Cells[1, 1, Rows, 1].Style.Numberformat.Format = "#,##0";
                        //ws.Cells[1, 3, Rows, 3].Style.Numberformat.Format = "YYYY-MM-DD";
                        //ws.Cells[1, 4, Rows, 5].Style.Numberformat.Format = "#,##0.00";

                        // add comp fields
                        // worksheet.Cell(5, 2).Formula = string.Format("SUM({0}:{1})", calcStartAddress, calcEndAddress); 

                        // Autofit
                        if (templateFile == null)
                        {
                            ws.Cells[start - 1, stlpec, riadok, activeColCount + stlpec - 1].AutoFitColumns();
                            foreach (var row in rowConfig.Where(o => o.order != 0))
                                if (ws.Column(row.order).Width < row.minsize)
                                    ws.Column(row.order).Width = row.minsize;
                        }

                        //Create an autofilter(global settings) for the range                           
                        if (def.autofilter)
                            ws.Cells[start - 1, stlpec, riadok, activeColCount + stlpec - 1].AutoFilter = true;
                    }

                    if (enabledTimestamp && def.timestamp)
                        ws.Cells[riadok + 2, stlpec].Value = "Created :  " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss");
                }

                package.Workbook.Calculate();

                // Set document properties
                package.Workbook.Properties.Comments = "Created with EpSqlGen Copyright © 2018 Miroslav Dubovský";
                package.Workbook.Properties.Created = DateTime.Now;
                package.Workbook.Properties.Title = outFileName;
                var pomProp = System.Configuration.ConfigurationManager.AppSettings.Get("Author");
                if (pomProp != null)
                    package.Workbook.Properties.Author = pomProp;
                pomProp = System.Configuration.ConfigurationManager.AppSettings.Get("Company");
                if (pomProp != null)
                    package.Workbook.Properties.Company = pomProp;
                if (def.version != null)
                    package.Workbook.Properties.Subject = "Template " + defFileName + ", version:" + def.version;
                else
                    package.Workbook.Properties.Subject = "Template " + defFileName;

                package.Save();
            }
            if (!enabledConsoleLog)
                Console.Write(JsonConvert.SerializeObject(new { output_file = newFile.FullName }, Formatting.Indented));
            else
                SimpleLog.WriteLog("Generated file: " + newFile.FullName);

        }
    }
}
