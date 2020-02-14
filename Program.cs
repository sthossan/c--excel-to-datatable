using Microsoft.VisualBasic.FileIO;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ConsoleApp
{
    public class Program
    {
        static string strFilePath = @"Book1.xlsx"; // location your csv or excel file
        static string extension = Path.GetExtension(strFilePath);
        static string connString = string.Empty;

        static DataTable dt;

        static void Main(string[] args)
        {
            DataTable dt = ConvertXSLXtoDataTable();
            DataTable dtTextFieldParser = ConvertCSVtoDataTableTextFieldParser();
            DataTable dtExcel = ConvertXSLXtoDataTable();
        }

        public static DataTable ConvertCSVtoDataTable()
        {
            dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    if (rows.Length > 0)
                    {
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i].Trim();
                        }
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        public static DataTable ConvertCSVtoDataTableTextFieldParser()
        {
            DataTable dt = new DataTable();

            using (TextFieldParser reader = new TextFieldParser(strFilePath))
            {
                reader.SetDelimiters(new string[] { "," });
                reader.HasFieldsEnclosedInQuotes = true;
                string[] colFields = reader.ReadFields();
                foreach (string column in colFields)
                {
                    DataColumn datecolumn = new DataColumn(column);
                    datecolumn.AllowDBNull = true;
                    dt.Columns.Add(datecolumn);
                }
                while (!reader.EndOfData)
                {
                    string[] fieldData = reader.ReadFields();
                    for (int i = 0; i < fieldData.Length; i++)
                    {
                        if (fieldData[i] == "")
                        {
                            fieldData[i] = null;
                        }
                    }
                    dt.Rows.Add(fieldData);
                }
            }

            return dt;
        }

        public static DataTable ConvertXSLXtoDataTable()
        {
            if (extension.Trim() == ".xls")
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";

            else if (extension.Trim() == ".xlsx")
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";

            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                oledbConn.Open();
                using (OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn))
                {
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    oleda.SelectCommand = cmd;
                    dt = new DataTable();
                    oleda.Fill(dt);
                    return dt;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oledbConn.Close();
            }
        }

    }

}
