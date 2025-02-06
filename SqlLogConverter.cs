using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using MySql.Data.MySqlClient;

namespace MeretFelugyelo.LogConverter
{
    internal class SqlLogConverter
    {
        string connectionString = "server=131.0.1.92;uid=robi;database=gerison_import_data";
        public static string sqlLogPath = @"A:\gerison_arlista\import_result.txt";
        public static string wizardPath = "C:\\Users\\LP-KATALOGUS1\\Desktop\\VS Munka\\MeretFelugyelo\\bin\\Debug\\import_wizard\\";
        public static string[] tableKeys = new string[] { "Hibak:", "Keszleten van, de nalunk nincs, hibara utal:", "Keszleten van,es nalunk megvan:" };
        public static Dictionary<string, SqlStockError> stockErrors = new Dictionary<string, SqlStockError>();
        public static Dictionary<string, SqlPriceError> priceErrors = new Dictionary<string, SqlPriceError>();
        public static Dictionary<string, string> inStockError = new Dictionary<string, string>();
        public static Dictionary<string, RightStock> inStock = new Dictionary<string, RightStock>();
        public static Dictionary<string, int> gyartoSor = new Dictionary<string, int>();
        public static int reportCol = 2;
        public static string lastString = "";

        public void ConvertFile(bool wizard = false)
        {
            ReadFile(wizard);
            Console.WriteLine("Stock Errors: " + stockErrors.Count);
            Console.WriteLine("Price Errors: " + priceErrors.Count);
            Console.WriteLine("Nalunk van 0: " + inStockError.Count);
            Console.WriteLine("Nalunk van 1: " + inStock.Count);
            Console.WriteLine("-----------------");
            if (!wizard) ConstructXmlFile();
            SaveToMySQL(wizard);
            if (!wizard) GenerateReport();
        }

        private int AddToNullReport(IXLWorksheet ws, MySqlDataReader reader, int row)
        {
            if (!gyartoSor.ContainsKey(reader[0].ToString()))
            { gyartoSor.Add(reader[0].ToString(), row); row++; }

            int insertRow = gyartoSor[reader[0].ToString()];
            if (lastString != reader[2].ToString())
                { reportCol++; lastString = reader[2].ToString(); }

            ws.Cell(insertRow, 1).Value = reader[0].ToString();
            ws.Cell(1, reportCol).Value = reader[2].ToString();
            ws.Cell(insertRow, reportCol).Value = reader[1].ToString();

            if (reportCol - 1 != 2 && ws.Cell(insertRow, reportCol - 1).Value.ToString() != reader[1].ToString())
                ws.Range(ws.Cell(insertRow, reportCol - 1), ws.Cell(insertRow, reportCol)).Style.Fill.BackgroundColor = XLColor.Amber;

            if (DateTime.Parse(reader[2].ToString()).Date == DateTime.Today.Date && reader[1].ToString() != "NULL")
                if (ws.Cell(insertRow, reportCol - 1).Value.ToString() != reader[1].ToString())
                    ws.Cell(insertRow, 2).Value = -(int.Parse(ws.Cell(insertRow, reportCol - 1).Value.ToString()) - int.Parse(reader[1].ToString()));

            return row;
        }

        private int AddToKeszletReport(IXLWorksheet ws, MySqlDataReader reader, int row)
        {
            if (!gyartoSor.ContainsKey(reader[0].ToString()))
            { gyartoSor.Add(reader[0].ToString(), row); row++; }

            int insertRow = gyartoSor[reader[0].ToString()];
            if (lastString != reader[3].ToString())
            { reportCol++; lastString = reader[3].ToString(); }

            ws.Cell(insertRow, 1).Value = reader[0].ToString();
            ws.Cell(1, reportCol).Value = reader[3].ToString();
            ws.Cell(insertRow, reportCol).Value = reader[1].ToString();

            if (reportCol - 1 != 2 && ws.Cell(insertRow, reportCol - 1).Value.ToString() != reader[1].ToString())
                ws.Range(ws.Cell(insertRow, reportCol - 1), ws.Cell(insertRow, reportCol)).Style.Fill.BackgroundColor = XLColor.Amber;

            if (DateTime.Parse(reader[3].ToString()).Date == DateTime.Today.Date && reader[1].ToString() != "NULL" && ws.Cell(insertRow, reportCol - 1).Value.ToString() != "NULL")
                if (ws.Cell(insertRow, reportCol - 1).Value.ToString() != reader[1].ToString())
                    ws.Cell(insertRow, 2).Value = -(int.Parse(ws.Cell(insertRow, reportCol - 1).Value.ToString()) - int.Parse(reader[1].ToString()));

            return row;
        }

        private int AddToOssszesReport(IXLWorksheet ws, MySqlDataReader reader, int row)
        {
            if (!gyartoSor.ContainsKey(reader[0].ToString()))
            { gyartoSor.Add(reader[0].ToString(), row); row++; }

            int insertRow = gyartoSor[reader[0].ToString()];
            if (lastString != reader[3].ToString())
            { reportCol++; lastString = reader[3].ToString(); }

            ws.Cell(insertRow, 1).Value = reader[0].ToString();
            ws.Cell(1, reportCol).Value = reader[3].ToString();
            ws.Cell(insertRow, reportCol).Value = reader[2].ToString();

            if (reportCol - 1 != 2 && ws.Cell(insertRow, reportCol - 1).Value.ToString() != reader[2].ToString())
                ws.Range(ws.Cell(insertRow, reportCol - 1), ws.Cell(insertRow, reportCol)).Style.Fill.BackgroundColor = XLColor.Amber;

            if (DateTime.Parse(reader[3].ToString()).Date == DateTime.Today.Date && reader[2].ToString() != "NULL" && ws.Cell(insertRow, reportCol - 1).Value.ToString() != "NULL")
                if (ws.Cell(insertRow, reportCol - 1).Value.ToString() != reader[2].ToString())
                    ws.Cell(insertRow, 2).Value = -(int.Parse(ws.Cell(insertRow, reportCol - 1).Value.ToString()) - int.Parse(reader[2].ToString()));

            return row;
        }

        public void GenerateReport()
        {
            //Null hiba
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("null_hiba");

            try
            {
                MySqlConnection mySqlConnection = new MySqlConnection(connectionString);
                mySqlConnection.Open();
                MySqlCommand command = mySqlConnection.CreateCommand();
                command.CommandText = "SELECT * FROM null_hiba h ORDER BY h.dateStamp";
                MySqlDataReader reader = command.ExecuteReader();
                int row = 2;
                while (reader.Read())
                {
                    row = AddToNullReport(ws, reader, row);
                }
            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex);
            }

            ws.SheetView.FreezeColumns(2);

            var wsKeszlet = wb.AddWorksheet("nalunk_van_keszletes");
            var wsOsszes = wb.AddWorksheet("nalunk_van_osszes");
            try
            {
                MySqlConnection mySqlConnection = new MySqlConnection(connectionString);
                mySqlConnection.Open();
                MySqlCommand command = mySqlConnection.CreateCommand();
                command.CommandText = "SELECT * FROM nalunk_van h ORDER BY h.dateStamp";
                MySqlDataReader reader = command.ExecuteReader();
                int row = 2;
                reportCol = 2;
                while (reader.Read())
                {
                    row = AddToKeszletReport(wsKeszlet, reader, row);
                    row = AddToOssszesReport(wsOsszes, reader, row);
                }
            }

            catch (MySqlException ex)
            {
                Console.WriteLine(ex);
            }

            try { wb.SaveAs("import_result_kimutatas.xlsx"); }
            catch (Exception e) { Console.WriteLine(e); }
        }

        void ReadFile(bool wizard = false)
        {
            StreamReader sr;
            if (!wizard) { sr = new StreamReader(sqlLogPath); }
            else
            {
                string fileName = Console.ReadLine();
                Console.WriteLine(File.Exists(wizardPath + fileName));
                if (File.Exists(wizardPath + fileName))
                    sr = new StreamReader(wizardPath + fileName);
                else
                    return;
            }

            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine();
                if (IsTableKey(line) == 0)
                    ReadErrors(sr, 0, 0);
                if (IsTableKey(line) == 1)
                    ReadStockError(sr, 0);
                if (IsTableKey(line) == 2)
                    ReadStockRight(sr, 0);
            }
        }

        void ReadStockRight(StreamReader sr, int tableCornerCount)
        {
            Console.WriteLine("*");
            if (tableCornerCount == 3)
                return;
            string line = sr.ReadLine();
            if (line == "")
            { ReadStockRight(sr, tableCornerCount); return; };
            line = line.Replace(" ", "");
            if (line[0] == '+')
            { tableCornerCount++; ReadStockRight(sr, tableCornerCount); return; }
            string[] data = line.Split('|');
            if (data[1] == "supplier")
            { ReadStockRight(sr, tableCornerCount); return; }

            RightStock stock = new RightStock();
            stock.keszlet = data[2];
            stock.osszes = data[3];

            inStock.Add(data[1], stock);

            ReadStockRight(sr, tableCornerCount);
        }

        int IsTableKey(string line)
        {
            if (tableKeys.Contains(line))
            {
                if (Array.IndexOf(tableKeys, line) == 0)
                    return 0;
                if (Array.IndexOf(tableKeys, line) == 1)
                    return 1;
                if (Array.IndexOf(tableKeys, line) == 2)
                    return 2;
            }
            return -1;
        }

        void ReadErrors(StreamReader sr, int tableCornerCount, int table)
        {
            if (tableCornerCount == 6)
                return;
            string line = sr.ReadLine();
            if (line == "")
                { ReadErrors(sr, tableCornerCount, table); return; };
            line = line.Replace(" ", "");
            if (line[0] == '+')
                { tableCornerCount++; ReadErrors(sr, tableCornerCount, table); return; }

            string[] data = line.Split('|');
            if (data[2] == "fresh_stock_hours" || data[2] == "fresh_price_hours")
                { table++; ReadErrors(sr, tableCornerCount, table); return; }

            if (table == 1)
            {
                SqlStockError error = new SqlStockError();
                error.id = data[1];
                error.freshStockHours = data[2];
                error.freshStockQuantity = data[3];
                error.freshStockCount = data[4];

                stockErrors.Add(data[1], error);
            }

            if (table == 2)
            {
                SqlPriceError error = new SqlPriceError();
                error.id = data[1];
                error.freshPriceHours = data[2];
                error.freshPriceQuantity = data[3];
                error.freshPriceCount = data[4];

                priceErrors.Add(data[1], error);
            }

            ReadErrors(sr, tableCornerCount, table);
        }

        void ReadStockError(StreamReader sr, int tableCornerCount)
        {
            if (tableCornerCount == 3)
                return;
            string line = sr.ReadLine();
            if (line == "")
                { ReadStockError(sr, tableCornerCount); return; };
            line = line.Replace(" ", "");
            if (line[0] == '+')
                { tableCornerCount++; ReadStockError(sr, tableCornerCount); return; }
            string[] data = line.Split('|');
            if (data[1] == "supplier")
                { ReadStockError(sr, tableCornerCount); return; }

            inStockError.Add(data[1], data[2]);

            ReadStockError(sr, tableCornerCount);
        }

        void ConstructXmlFile()
        {
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.AddWorksheet("StockHiba");

            int i = 2;
            foreach (SqlStockError item in stockErrors.Values)
            {
                ws.Cell(i, 1).Value = item.id;
                ws.Cell(i, 2).Value = item.freshStockHours;
                ws.Cell(i, 3).Value = item.freshStockQuantity;
                ws.Cell(i, 4).Value = item.freshStockCount;
                i++;
            }
            ws.Cell(1, 1).Value = "id";
            ws.Cell(1, 2).Value = "fresh_stock_hours";
            ws.Cell(1, 3).Value = "fresh_stock_quantity";
            ws.Cell(1, 4).Value = "fresh_stock_count";
            ws = wb.AddWorksheet("PriceHiba");
            i = 2;
            foreach (SqlPriceError item in priceErrors.Values)
            {
                ws.Cell(i, 1).Value = item.id;
                ws.Cell(i, 2).Value = item.freshPriceHours;
                ws.Cell(i, 3).Value = item.freshPriceQuantity;
                ws.Cell(i, 4).Value = item.freshPriceCount;
                i++;
            }
            ws.Cell(1, 1).Value = "id";
            ws.Cell(1, 2).Value = "fresh_price_hours";
            ws.Cell(1, 3).Value = "fresh_price_quantity";
            ws.Cell(1, 4).Value = "fresh_price_count";
            ws = wb.AddWorksheet("NullaHiba");
            i = 2;
            foreach (var item in inStockError)
            {
                ws.Cell(i, 1).Value = item.Key;
                ws.Cell(i, 2).Value = item.Value;
                i++;
            }
            ws.Cell(1, 1).Value = "supplier";
            ws.Cell(1, 2).Value = "count";
            try { wb.SaveAs("import_result.xlsx"); }
            catch (Exception e) { Console.WriteLine(e); }
        }

        void SaveToMySQL(bool wizard = false)
        {
            string dateString = "";
            if (wizard)
                dateString = Console.ReadLine();
            try
            {
                MySqlConnection mySqlConnection = new MySqlConnection(connectionString);
                mySqlConnection.Open();

                foreach (var item in inStockError)
                {
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = mySqlConnection;
                    command.CommandText = "INSERT INTO null_hiba(name,value,dateStamp) VALUES(@name,@value,@dateStamp)";
                    command.Parameters.AddWithValue("@name", item.Key);
                    command.Parameters.AddWithValue("@value", item.Value);
                    command.Parameters.AddWithValue("@dateStamp", (wizard) ? dateString : DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"));
                    command.ExecuteNonQuery();
                }

                foreach (SqlStockError item in stockErrors.Values)
                {
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = mySqlConnection;
                    command.CommandText = "INSERT INTO stock_hiba(name,fresh_stock_hours,fresh_stock_quantity,fresh_stock_count,dateStamp) VALUES(@name,@fresh_stock_hours,@fresh_stock_quantity,@fresh_stock_count,@dateStamp) ";
                    command.Parameters.AddWithValue("@name", item.id);
                    command.Parameters.AddWithValue("@fresh_stock_hours", item.freshStockHours);
                    command.Parameters.AddWithValue("@fresh_stock_quantity", item.freshStockQuantity);
                    command.Parameters.AddWithValue("@fresh_stock_count", item.freshStockCount);
                    command.Parameters.AddWithValue("@dateStamp", (wizard) ? dateString : DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"));
                    command.ExecuteNonQuery();
                }

                foreach (SqlPriceError item in priceErrors.Values)
                {
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = mySqlConnection;
                    command.CommandText = "INSERT INTO price_hiba(name,fresh_price_hours,fresh_price_quantity,fresh_price_count,dateStamp) VALUES(@name,@fresh_price_hours,@fresh_price_quantity,@fresh_price_count,@dateStamp)";
                    command.Parameters.AddWithValue("@name", item.id);
                    command.Parameters.AddWithValue("@fresh_price_hours", item.freshPriceHours);
                    command.Parameters.AddWithValue("@fresh_price_quantity", item.freshPriceQuantity);
                    command.Parameters.AddWithValue("@fresh_price_count", item.freshPriceCount);
                    command.Parameters.AddWithValue("@dateStamp", (wizard) ? dateString : DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"));
                    command.ExecuteNonQuery();
                }

                foreach (var item in inStock)
                {
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = mySqlConnection;
                    command.CommandText = "INSERT INTO nalunk_van(name,keszletes,osszes,dateStamp) VALUES(@name,@keszletes,@osszes,@dateStamp)";
                    command.Parameters.AddWithValue("@name", item.Key);
                    command.Parameters.AddWithValue("@keszletes", item.Value.keszlet);
                    command.Parameters.AddWithValue("@osszes", item.Value.osszes);
                    command.Parameters.AddWithValue("@dateStamp", (wizard) ? dateString : DateTime.Today.ToString("yyyy-MM-dd HH:mm:ss"));
                    command.ExecuteNonQuery();
                }
                mySqlConnection.Close();
            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex);
            }
        }
    }

    public struct RightStock
    {
        public string keszlet { get; set; }
        public string osszes { get; set; }
    }

    public struct SqlStockError
    {
        public string id { get; set; }
        public string freshStockHours { get; set; }
        public string freshStockQuantity { get; set; }
        public string freshStockCount { get; set; }

        public SqlStockError(string id, string freshStockHours, string freshStockQuantity, string freshStockCount)
        {
            this.id = id;
            this.freshStockHours = freshStockHours;
            this.freshStockQuantity = freshStockQuantity;
            this.freshStockCount = freshStockCount;
        }
    }

    public struct SqlPriceError
    {
        public string id { get; set; }
        public string freshPriceHours { get; set; }
        public string freshPriceQuantity { get; set; }
        public string freshPriceCount { get; set; }

        public SqlPriceError(string id, string freshPriceHours, string freshPriceQuantity, string freshPriceCount)
        {
            this.id = id;
            this.freshPriceHours = freshPriceHours;
            this.freshPriceQuantity = freshPriceQuantity;
            this.freshPriceCount = freshPriceCount;
        }
    }
}
