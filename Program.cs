using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace SQLtoExcel
{
    class Program
    {
        static DirectoryInfo ScriptsFolder;
        static SqlConnectionStringBuilder SqlConnection;
        static FileInfo ExcelSpreadsheet;
        static DataSet AllDatatablesForExcel = new DataSet();
        static bool Help = false;

        static void Main(string[] args)
        {
            try
            {
                GetCommandLineParams(args);

                if (Help)
                {
                    ShowHelp();
                }
                else
                {
                    PopulateDataSet();
                    CreateExcelSpreadsheet();
                }
            }
            catch(Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine(ex.ToString());
            }

            Console.WriteLine("");
            Console.WriteLine("Enter any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Self-explanatory.
        /// </summary>
        private static void ShowHelp()
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);
            Console.WriteLine("Saves the output from one or more SQL scripts to an Excel spreadsheet file.");
            Console.WriteLine(Environment.NewLine);

            Console.WriteLine("Command line parameters:");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\t/Server");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\t\tRequired: name of the SQL Server instance.");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\t/Database");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\tRequired: name of the database.");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\t/Login");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\t\tOptional: login name for SQL Authentication (omit for Windows Authentication).");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\t/Password");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\tOptional: password for SQL Authentication (omit for Windows Authentication).");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\t/ScriptsFolder");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\tOptional: path to folder containing *.sql scripts (defaults to executable path).");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\t/ExcelFile");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\tOptional: path of Excel spreadsheet file to be created (defaults to \"" +
                System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".xlsx\" in executable path).");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\t/Help");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write(" or ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("/?");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\tOptional: displays this screen.");


            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Commandline parameter usage: /Parameter=value");
            Console.WriteLine("Example:\t\t /Server=" + Environment.MachineName + " /Database=master");
        }

        /// <summary>
        /// Populates the AllDatatablesForExcel dataset.
        /// </summary>
        private static void PopulateDataSet()
        {
            using (SqlConnection conn = new SqlConnection(SqlConnection.ConnectionString))
            {
                conn.Open();

                #region Iterate through SQL scripts
                foreach (FileInfo fi in ScriptsFolder.GetFiles("*.sql"))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(fi.Name);

                    foreach (string batch in GetBatches(fi))
                    {
                        if (string.IsNullOrEmpty(batch.Trim()))
                        {
                            continue;
                        }

                        using (SqlCommand cmd = new SqlCommand())
                        {
                            cmd.Connection = conn;
                            cmd.CommandType = System.Data.CommandType.Text;
                            cmd.CommandText = batch;
                            cmd.CommandTimeout = 0;

                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                using (System.Data.DataSet ds = new System.Data.DataSet())
                                {
                                    try
                                    {
                                        da.Fill(ds);
                                    }
                                    catch
                                    {
                                        Console.ForegroundColor = ConsoleColor.Cyan;
                                        Console.Write("\t");
                                        Console.WriteLine("Script failed: " + fi.Name);
                                    }

                                    foreach (DataTable dt in ds.Tables)
                                    {
                                        DataTable dtCopy = dt.Copy();
                                        dtCopy.TableName = fi.Name.Replace(fi.Extension, "");

                                        dtCopy.TableName = dtCopy.TableName.Substring(0, dtCopy.TableName.Length > 31 ? 31 : dtCopy.TableName.Length);
                                        AllDatatablesForExcel.Tables.Add(dtCopy);
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
            }
        }

        /// <summary>
        /// Returns a string list of T-SQL batches from the input FileInfo.
        /// </summary>
        /// <param name="fi">Represents a *.sql file.</param>
        /// <returns></returns>
        private static List<string> GetBatches(FileInfo fi)
        {
            List<string> batches = new List<string>();
            StringBuilder sb = new StringBuilder();

            //read all lines of the file.
            string [] fileLines = File.ReadAllLines(fi.FullName);

            //iterate through the file lines.
            for(int i = 0; i < fileLines.Length; i++)
            {
                //search for "GO" batch separators.
                if(fileLines[i].Trim().ToUpper() == "GO".ToUpper())
                {
                    //When "GO" is found, add the contents of the
                    //StringBuilder to the list of batches.
                    if(!string.IsNullOrEmpty(sb.ToString().Trim()))
                    {
                        batches.Add(sb.ToString());
                    }

                    sb.Clear();
                }
                else
                {
                    //If it's not a "GO" batch separator,
                    //add the file line to the StringBuilder.
                    sb.AppendLine(fileLines[i]);
                }
            }

            if (!string.IsNullOrEmpty(sb.ToString().Trim()))
            {
                batches.Add(sb.ToString());
            }

            sb.Clear();

            return batches;
        }

        /// <summary>
        /// Creates a Microsoft Excel Worksheet file using data from the AllDatatablesForExcel dataset.
        /// </summary>
        private static void CreateExcelSpreadsheet()
        {
            using (MemoryStream ms = Utils.ExportDataSetToExcel(AllDatatablesForExcel))
            {
                using (FileStream fs = new FileStream(ExcelSpreadsheet.FullName, FileMode.Create))
                {
                    ms.WriteTo(fs);
                }
            }
        }

        /// <summary>
        /// Parses command line arguments, gathers parameter values.
        /// </summary>
        /// <param name="args"></param>
        private static void GetCommandLineParams(string[] args)
        {
            
            if(args.Contains("/Help", StringComparer.OrdinalIgnoreCase) || args.Contains("/?"))
            {
                Help = true;
                return;
            }

            SqlConnection = new SqlConnectionStringBuilder();

            foreach (string arg in args)
            {
                string[] argParts = arg.Split('=');

                if (argParts.Length == 2)
                {
                    string param = argParts[0];
                    string value = argParts[1].Replace("\"", "");

                    if (string.Compare(param, "/Server", true) == 0)
                    {
                        SqlConnection.DataSource = value;
                    }
                    else if (string.Compare(param, "/Database", true) == 0)
                    {
                        SqlConnection.InitialCatalog = value;
                    }
                    else if (string.Compare(param, "/Login", true) == 0)
                    {
                        SqlConnection.UserID = value;
                    }
                    else if (string.Compare(param, "/Password", true) == 0)
                    {
                        SqlConnection.Password = value;
                    }
                    else if (string.Compare(param, "/ScriptsFolder", true) == 0)
                    {
                        ScriptsFolder = new DirectoryInfo(value);
                    }
                    else if (string.Compare(param, "/ExcelFile", true) == 0)
                    {
                        ExcelSpreadsheet = new FileInfo(value);
                    }
                }
            }

            if(ScriptsFolder == null || !ScriptsFolder.Exists)
            {
                //TODO: do something here.
                ScriptsFolder = new DirectoryInfo(System.Reflection.Assembly.GetEntryAssembly().Location);
            }
            
            if(string.IsNullOrEmpty(SqlConnection.UserID))
            {
                //If no login provided, assume integrated security.
                SqlConnection.IntegratedSecurity = true;
            }

            if (ExcelSpreadsheet == null)
            {
                //create spreadsheet in executable path.
                ExcelSpreadsheet = new FileInfo(Path.Combine(System.Reflection.Assembly.GetEntryAssembly().Location,
                    System.Reflection.Assembly.GetExecutingAssembly().GetName().Name) + ".xlsx");
            }

            //Console.WriteLine("Scripts Folder: " + ScriptsFolder.FullName);

        }
    }
}
