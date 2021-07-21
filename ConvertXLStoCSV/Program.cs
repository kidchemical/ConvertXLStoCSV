using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ConvertXLStoCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.ForegroundColor = ConsoleColor.White;
            string inputPath = @"C:\Users\User\Desktop\History.xls";
            string outputPath = @"C:\Users\User\Desktop\History.csv";
            int worksheetNumber = 1;

            if (args.Length > 1)
            {
                try
                {
                    if (args[0] != null) inputPath = args[0];
                    if (args[1] != null) outputPath = args[1];
                }
                catch
                {
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Unhandled exception has occured. Exiting...");
                    throw;
                }
            }


            try
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue;
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("XLS to CSV Converter - Jeremy Pudget, Ian McCambridge 2021 - Wise Auto Group IT");
                Console.WriteLine("-------------------------------------------------------------------------------");
                Console.WriteLine("[Argument1] (XLS) Input path: " + inputPath);
                Console.WriteLine("[Argument2] (CSV) Output path: " + outputPath);
                try
                {
                    worksheetNumber = Int32.Parse(args[2]);
                    Console.WriteLine("[Argument3] Worksheet: " + worksheetNumber.ToString());
                }
                catch
                {
                    Console.BackgroundColor = ConsoleColor.DarkBlue;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[Argument3] No worksheet defined. Converting default (1) worksheet...");
                    Console.BackgroundColor = ConsoleColor.DarkBlue;
                    Console.ForegroundColor = ConsoleColor.White;
                }
                Console.WriteLine("-------------------------------------------------------------------------------");

                Console.BackgroundColor = ConsoleColor.Black;
                Console.ForegroundColor = ConsoleColor.White;
                ConvertExcelToCsv(inputPath, outputPath, worksheetNumber);
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.Black;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("An error has occured. (Perhaps your path arguments are incorrect or the CSV output file exists).");
                throw e;
            }
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Completed with success!");
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.White;

            static void ConvertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
            {
                //Console.WriteLine("DEBUG:::" + csvOutputFile);
                if (!File.Exists(excelFilePath)) throw new  FileNotFoundException(excelFilePath);
                if (File.Exists(csvOutputFile))
                {
                    Console.WriteLine("CSV file already exists! Overwriting...");
                    File.Delete(csvOutputFile);
                    //throw new ArgumentException("File exists: " + csvOutputFile);
                }

                // connection string
                var cnnStr = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", excelFilePath);
                var cnn = new OleDbConnection(cnnStr);

                // get schema, then data
                var dt = new DataTable();
                try
                {
                    cnn.Open();
                    var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                    string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                    string sql = String.Format("select * from [{0}]", worksheet);
                    var da = new OleDbDataAdapter(sql, cnn);
                    da.Fill(dt);
                }
                catch (Exception e)
                {

                    Console.BackgroundColor = ConsoleColor.DarkBlue;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Unhandled exception has occured. Exiting...");
                    throw e;
                }
                finally
                {
                    // free resources
                    cnn.Close();
                }

                // write out CSV data
                using (var wtr = new StreamWriter(csvOutputFile))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        bool firstLine = true;
                        foreach (DataColumn col in dt.Columns)
                        {
                            if (!firstLine) { wtr.Write(","); } else { firstLine = false; }
                            var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                            wtr.Write(String.Format("\"{0}\"", data));
                        }
                        wtr.WriteLine();
                    }
                }
            }
        }
    }

}
