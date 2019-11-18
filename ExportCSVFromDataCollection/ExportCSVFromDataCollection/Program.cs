using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportCSVFromDataCollection
{
    class Program
    {
        
        static void Main(string[] args)
        {
            // INPUT
            var pathToLoadCSV = @"C:\Users\adm.m.gosciniak\Desktop\csv.csv";
            var pathToSaveCSV = @"C:\Users\adm.m.gosciniak\Desktop\csv2.csv";
            var sourceTextQualifier = ",";
            var firstLineIsHeader = true;
            // OUTPUT
            var dataTable = new DataTable();
            var success = true;
            var message = string.Empty;

            ImportCSVtoDataTable(pathToLoadCSV, ref dataTable, ref message, ref success, sourceTextQualifier, firstLineIsHeader);
            PrintInConsoleDataTable(dataTable);
            ExportCSV(dataTable, pathToSaveCSV, ref message, ref success, sourceTextQualifier);
            Console.ReadKey();
        }
        static void PrintInConsoleDataTable(DataTable dt)
        {
            var c = string.Empty;
            foreach (DataColumn col in dt.Columns)
                c = c + col.ColumnName + ";";
            c = c.Remove(c.Length - 1, 1);
            Console.WriteLine(c);

            var r = string.Empty;
            foreach (DataRow row in dt.Rows)
                Console.WriteLine(string.Join(";", row.ItemArray.Select(i => "" + i)));
        }
        private static void ImportCSVtoDataTable(string pathToLoadCSV, ref DataTable dataCollection, ref string message, ref bool success, string sourceTextQualifier = ";", bool firstLineIsHeader = true)
        {
            try
            {
                if (!File.Exists(pathToLoadCSV))
                    throw new FileNotFoundException();

                var csvText = File.ReadAllLines(pathToLoadCSV).ToList();
                var separator = sourceTextQualifier.ToCharArray();

                if (firstLineIsHeader)
                {
                    var columns = csvText[0].Split(separator);

                    foreach (var field in columns)
                    {
                        var col = new DataColumn(field, typeof(string));
                        dataCollection.Columns.Add(col);
                    }
                }

                csvText.RemoveAt(0);

                foreach (var line in csvText)
                {
                    var fields = line.Split(separator);
                    dataCollection.Rows.Add(fields);
                }

                success = true;
                message = "";
            }
            catch (Exception e)
            {
                success = false;
                message = e.Message;
            }
        }
        private static void ImportCSVtoDataSetByOLEDB(string pathToLoadCSV, ref DataSet dataSet, ref string message, ref bool success, bool First_Line_Is_Header = true)
        {
            try
            {
                string HDRString = "No";
                if (First_Line_Is_Header)
                    HDRString = "Yes";

                if (!File.Exists(pathToLoadCSV))
                    throw new FileNotFoundException();

                var FileName = Path.GetFileName(pathToLoadCSV);
                var Folder = Path.GetDirectoryName(pathToLoadCSV);

                var cn = new OleDbConnection($@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={Folder};Extended Properties=""Text;HDR={HDRString};FMT=Delimited""");
                var da = new OleDbDataAdapter();
                var ds = new DataSet();
                var cd = new OleDbCommand($"SELECT * FROM [{FileName}]", cn);

                cn.Open();
                da.SelectCommand = cd;
                ds.Clear();
                da.Fill(ds, "CSV");
                dataSet = ds;
                cn.Close();

                success = true;
                message = "";
            }
            catch (Exception e)
            {
                success = false;
                message = e.Message;
            }
        }
        private static void ImportCSVByOLEDB(string CSV_Location, bool First_Line_Is_Header, ref DataSet Values) //Coverted to c# from Blueprism
        {
            string HDRString = "No";
            if (First_Line_Is_Header)
                HDRString = "Yes";

            if (!File.Exists(CSV_Location))
                throw new FileNotFoundException();

            var FileName = Path.GetFileName(CSV_Location);
            var Folder = Path.GetDirectoryName(CSV_Location);

            var cn = new OleDbConnection($@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={Folder};Extended Properties=""Text;HDR={HDRString};FMT=Delimited""");
            var da = new OleDbDataAdapter();
            var ds = new DataSet();
            var cd = new OleDbCommand($"SELECT * FROM [{FileName}]", cn);

            cn.Open();
            da.SelectCommand = cd;
            ds.Clear();
            da.Fill(ds, "CSV");
            Values = ds;
            cn.Close();
        }
        private static void ExportCSV(DataTable dataCollection, string pathToSaveCSV, ref string message, ref bool success, string sourceTextQualifier = ",", bool firstLineIsHeader = true)
        {
            try
            {
                using (StreamWriter sw = File.CreateText(pathToSaveCSV))
                {
                    if (firstLineIsHeader)
                    {
                        string c = string.Empty;
                        foreach (DataColumn col in dataCollection.Columns)
                            c = c + col.ColumnName + sourceTextQualifier;
                        c = c.Remove(c.Length - 1, 1);

                        sw.WriteLine(c.Trim());
                    }

                    string r = string.Empty;
                    foreach (DataRow row in dataCollection.Rows)
                        sw.WriteLine(string.Join(sourceTextQualifier, row.ItemArray.Select(i => "" + i)));
                }

                success = true;
                message = "";
            }
            catch (Exception e)
            {
                success = false;
                message = e.Message;
            }
        }
    }
}
