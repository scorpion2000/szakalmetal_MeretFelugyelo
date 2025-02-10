using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MeretFelugyelo.Analyzer
{
    internal class FileAnalyzer
    {
        public List<string> importantFiles = new List<string>();
        public string filePath = "\\\\Fs\\ARLISTA\\gerison_arlista\\";
        static string appPath = AppDomain.CurrentDomain.BaseDirectory;
        public Dictionary<string, FileParseData> parseData = new Dictionary<string, FileParseData>();
        public Dictionary<string, List<FileData>> fileData = new Dictionary<string, List<FileData>>();
        string[] titles = new string[] { "Kódnév", "Ár" };

        public FileAnalyzer(bool testData = false)
        {
            if (testData)
            {
                filePath = appPath + "\\TestDatas\\";
            }
        }

        public void CompareFileData()
        {
            XLWorkbook wb = new XLWorkbook();
            wb.AddWorksheet("main");

            GatherAllFileData();
            List<FileData> oldData = new List<FileData>();
            List<string> files = new List<string>();
            foreach (var item in fileData)
            {
                List<FileData> newData = new List<FileData>();
                List<FileData> deletedData = new List<FileData>();
                List<FileData> modifiedData = new List<FileData>();

                if (File.Exists(appPath + "\\save\\" + item.Key + ".json"))
                    oldData = JsonSerializer.Deserialize<List<FileData>>(File.ReadAllText(appPath + "\\save\\" + item.Key + ".json"));
                else { Console.WriteLine("New file: {0}", item.Key); continue; }
                Dictionary<string, FileData> oldDataIndexed = new Dictionary<string, FileData>();
                //Dictionary<string, FileData> removedIndexed = oldDataIndexed.ToDictionary(e => e.Key, e => e.Value);    //dumb
                Dictionary<string, FileData> removedIndexed = new Dictionary<string, FileData>();
                foreach (FileData keksz in oldData)
                {
                    if (keksz.strings.Count == 0) continue;
                    if (!oldDataIndexed.ContainsKey(keksz.strings[0]))
                    {
                        oldDataIndexed.Add(keksz.strings[0], keksz);
                        removedIndexed.Add(keksz.strings[0], keksz);
                    }
                    else
                        Console.WriteLine("Kulcs duplikáció a {0} fájlban: {1}", item.Key, keksz.strings[0]);
                }


                Console.WriteLine("Evaluating {0}", item.Key);
                files.Add(item.Key);
                foreach (FileData code in item.Value)
                {
                    if (code.strings.Count == 0) continue;
                    if (!oldDataIndexed.ContainsKey(code.strings[0]))
                        { newData.Add(code); continue; }

                    for ( int y = 0; y < code.strings.Count; y++)
                    {
                        if (oldDataIndexed[code.strings[0]].strings[y] != code.strings[y])
                            modifiedData.Add(code);
                    }
                    removedIndexed.Remove(code.strings[0]);
                }

                foreach (var code in removedIndexed)
                    deletedData.Add(code.Value);

                if (newData.Count == 0 && modifiedData.Count == 0 && deletedData.Count == 0)
                    continue;
                string sheetName = item.Key.Substring(0, Math.Min(item.Key.Length, 30));
                var ws = wb.AddWorksheet(sheetName);

                //Insert titles
                for (int col = 1; col < titles.Length; col++)
                    ws.Cell(1, col + 1).Value = titles[col];

                //Insert "old" titles
                for (int col = 1; col < titles.Length; col++)
                    ws.Cell(1, col + titles.Length + 1).Value = "Régi " + titles[col];

                int i = 2;
                i = GenerateFileValues(newData, item.Key, ws, "Új", i);
                i = GenerateFileValues(modifiedData, item.Key, ws, "Módosult", i, oldDataIndexed);
                i = GenerateFileValues(deletedData, item.Key, ws, "Törölt", i);
                ws.Column(titles.Length + 1).SetAutoFilter();
            }

            try { wb.SaveAs(appPath + "Rapid ár és cikkszám módosulás.xlsx"); }
            catch (Exception e) { Console.WriteLine(e); }
            CreateSaveFiles();
        }

        private void CreateSaveFiles()
        {
            var json = JsonSerializer.Serialize(parseData);
            File.WriteAllText(appPath + "fileDataSave_PARSE.json", json);
            foreach (var data in fileData)
            {
                json = JsonSerializer.Serialize(data.Value);
                File.WriteAllText(appPath + "\\save\\" + data.Key + ".json", json);
            }
        }

        private int GenerateFileValues(List<FileData> newData, string name, IXLWorksheet ws, string type, int i, Dictionary<string, FileData> oldDataIndexed = null)
        {
            foreach (var data in newData)
            {
                if (i > 1000000) break;
                int x = 1;
                string id = data.strings[0];
                foreach (string col in data.strings)
                {
                    ws.Cell(i, x).Value = col;
                    if (oldDataIndexed != null && oldDataIndexed.ContainsKey(id))
                        ws.Cell(i, data.strings.Count + x).Value = oldDataIndexed[id].strings[x - 1];
                    x++;
                }
                ws.Cell(i, data.strings.Count + 1).Value = type;
                i++;
            }
            return i;
        }

        public void GatherAllFileData()
        {
            Console.WriteLine("Gathering file datas...");
            List<string> validFiles = GatherValidFiles();
            if (File.Exists(appPath + "fileDataSave_PARSE.json"))
                parseData = JsonSerializer.Deserialize<Dictionary<string, FileParseData>>(File.ReadAllText(appPath + "fileDataSave_PARSE.json"));

            Console.WriteLine(validFiles.Count);
            foreach (string file in validFiles)
            {
                if (new FileInfo(file).Length == 0) continue;
                if (file == "VWStock.csv") continue;
                string extension = Path.GetExtension(filePath + Path.GetFileName(file));
                List<FileData> data = new List<FileData>();
                if (extension == ".xlsx")
                    data = GetDataFromXlsx(filePath, Path.GetFileName(file));
                if (extension == ".csv")
                    data = GetDataFromFile(filePath, Path.GetFileName(file), false);
                if (extension == ".txt")
                    data = GetDataFromFile(filePath, Path.GetFileName(file), true);
                Console.WriteLine(data.Count);
                fileData.Add(Path.GetFileName(file), data);
            }
        }

        private List<FileData> GetDataFromFile(string filePath, string fileName, bool askSeperator)
        {
            List<int> dataColumns = (parseData.ContainsKey(fileName) ? parseData[fileName].dataColumns : new List<int>());

            if (dataColumns.Count == 0)
            {
                FileParseData data = new FileParseData();
                Console.WriteLine("Fájl még nem volt mentve! {0}", fileName);
                Console.WriteLine("Van fejléc? (Y/N)");
                data.header = (Console.ReadLine().ToLower() == "y") ? true : false;

                data.dataColumns = GenerateDataColumns();
                if (askSeperator)
                {
                    Console.WriteLine("Tab az adatelválasztó karaker? (Y/N)");
                    if (Console.ReadLine().ToLower() == "y")
                    {
                        data.seperator = '\t';
                    }
                    else
                    {
                        Console.WriteLine("Mi az adatelválasztó karaker?");
                        data.seperator = Char.Parse(Console.ReadLine());
                    }
                }
                else
                {
                    data.seperator = ';';
                    Console.WriteLine("Első cellában van az adat? (Y/N)");
                    if (Console.ReadLine().ToLower() == "y")
                    {
                        data.firstCellData = true;
                        Console.WriteLine("Tab az első cella adatelválassztója? (Y/N)");
                        if (Console.ReadLine().ToLower() == "y")
                        {
                            data.firstCellDataSeperator = '\t';
                        }
                        else
                        {
                            Console.WriteLine("Mi az első cella adatelválasztó karaktere?");
                            data.firstCellDataSeperator = Char.Parse(Console.ReadLine());
                        }

                    }
                }
                parseData.Add(fileName, data);

                dataColumns = data.dataColumns;
            }
            List<FileData> newFileDatas = new List<FileData>();
            GetFileRowData(newFileDatas, filePath + fileName, dataColumns, parseData[fileName]);
            return newFileDatas;
        }

        private void GetFileRowData(List<FileData> fileDatas, string file, List<int> dataColumns, FileParseData parseData)
        {
            Console.WriteLine(file);
            StreamReader sr = new StreamReader(file);
            if (parseData.header) sr.ReadLine();
            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine();
                string[] lineData = line.Split(parseData.seperator);
                if (parseData.firstCellData)
                    lineData = lineData[0].Split(parseData.firstCellDataSeperator);

                FileData fileData = new FileData(new List<string>());
                foreach (int col in dataColumns)
                {
                    if (col >= lineData.Length)
                        continue;
                    fileData.strings.Add(lineData[col - 1]);
                }
                fileDatas.Add(fileData);
            }
        }

        private List<FileData> GetDataFromXlsx(string filePath, string fileName)
        {
            Console.WriteLine("Van fejléc? (Y/N)");
            bool header = (Console.ReadLine() == "Y") ? true : false;
            List<int> dataColumns = (parseData.ContainsKey(fileName) ? parseData[fileName].dataColumns : new List<int>());
            var workbook = new XLWorkbook(filePath + fileName);
            var worksheet = workbook.Worksheet(0);

            if (dataColumns.Count == 0)
            {
                FileParseData data = new FileParseData();
                data.dataColumns = GenerateDataColumns();
                data.seperator = '-';
                parseData.Add(fileName, data);

                dataColumns = data.dataColumns;
            }

            List<FileData> newFileDatas = new List<FileData>();
            GetXlsxRowData(newFileDatas, worksheet, dataColumns, (header) ? 1 : 2);
            return newFileDatas;
        }

        private void GetXlsxRowData(List<FileData> fileDatas, IXLWorksheet worksheet, List<int> dataColumns, int row)
        {
            var workRow = worksheet.Row(row);
            if (workRow.IsEmpty())
                return;

            FileData fileData = new FileData(new List<string>());
            foreach (int col in dataColumns)
                fileData.strings.Add(workRow.Cell(col).Value.ToString());
            fileDatas.Add(fileData);

            GetXlsxRowData(fileDatas, worksheet, dataColumns, row + 1);
        }

        private List<int> GenerateDataColumns()
        {
            List<int> ints = new List<int>();
            string response = "";
            Console.WriteLine("A fájlhoz nem tartozik táblázati adat");

            Console.WriteLine("Hányadik oszlopban találjuk a cikkszámot? (EXCEL -> A oszlop = 1)");
            ints.Add(Int32.Parse(Console.ReadLine()));

            Console.WriteLine("Hányadik oszlopban találjuk az ár értéket? (EXCEL -> A oszlop = 1) (-1 HA NINCS)");
            response = Console.ReadLine();
            if (response != "-1")
                ints.Add(Int32.Parse(response));

            /*Console.WriteLine("Hányadik oszlopban találjuk a súly értéket? (EXCEL -> A oszlop = 1) (-1 HA NINCS)");
            response = Console.ReadLine();
            if (response != "-1")
                ints.Add(Int32.Parse(response));*/

            return ints;
        }

        public void GatherImportantFiles()
        {
            Console.WriteLine("Gathering important files...");
            StreamReader sr = new StreamReader(appPath + "fajl_beszall.csv");
            sr.ReadLine();
            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine();
                string[] data = line.Split(';');
                if (data[6] == "TRUE")
                    importantFiles.Add(data[0]);
            }
        }

        public List<string> GatherValidFiles()
        {
            Console.WriteLine("Gathering valid files...");
            GatherImportantFiles();
            List<string> allFiles = Directory.GetFiles(filePath).ToList();
            List<string> validFiles = new List<string>();
            foreach (string file in allFiles)
            {
                string ext = Path.GetExtension(file);
                string name = Path.GetFileName(file);
                if (!importantFiles.Contains(name))
                    continue;
                if (ext == ".csv" || ext == ".txt" || ext == ".xlsx")
                    if (!name.Contains("zip"))
                        validFiles.Add(file);
            }
            return validFiles;
        }
    }

    public struct FileParseData
    {
        public char seperator { get; set; }
        public List<int> dataColumns { get; set; }
        public bool header { get; set; }
        public bool firstCellData { get; set; }
        public char firstCellDataSeperator { get; set; }

        public FileParseData(char seperator, List<int> dataColumns, bool header, bool firstCellData, char firstCellDataSeperator)
        {
            this.seperator = seperator;
            this.dataColumns = dataColumns;
            this.header = header;
            this.firstCellData = firstCellData;
            this.firstCellDataSeperator = firstCellDataSeperator;
        }
    }

    public struct FileData
    {
        public List<string> strings { get; set; }

        public FileData(List<string> strings)
        {
            this.strings = strings;
        }
    }
}