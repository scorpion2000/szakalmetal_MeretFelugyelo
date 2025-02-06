using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using System.Text.Json;
using Quartz;
using SMBLibrary;
using SMBLibrary.Client;
using System.Net;
using MeretFelugyelo.Analyzer;
using MeretFelugyelo.LogConverter;

namespace MeretFelugyelo
{
    internal class Program
    {
        static string filePath = "\\\\Fs\\ARLISTA\\gerison_arlista\\";
        static string appPath = "C:\\Users\\LP-KATALOGUS1\\Desktop\\VS Munka\\MeretFelugyelo\\bin\\Debug\\";
        static string savePath = appPath;
        static DateTime compareTime = new DateTime(2024, 10, 1, 17, 00, 0);
        static DateTime compareTimeTo = new DateTime(2024, 10, 1, 19, 45, 0);
        static Dictionary<string, LongDataSave> savedFilesLong;
        static Dictionary<string, SheetData> savedFileTitles;
        static void Main(string[] args)
        {
            SqlLogConverter logger = new SqlLogConverter();
            logger.ConvertFile();
            SaveToSmb("import_result.xlsx", "FGergo\\");
            SaveToSmb("import_result_kimutatas.xlsx", "FGergo\\");
            /*FileAnalyzer fileAnalyzer = new FileAnalyzer();
            fileAnalyzer.CompareFileData();
            //SaveToSmb("Rapid ár és cikkszám módosulás.xlsx", "FGergo\\");

            //Everything after this was made in a haste. It's.. not great..
            //I pray for the coding gods for forgiveness
            bool generateComparison = false;
            Dictionary<string, FileData> savedFiles = new Dictionary<string, FileData>();
            savedFilesLong = new Dictionary<string, LongDataSave>();
            savedFileTitles = new Dictionary<string, SheetData>();
            if (File.Exists(appPath + "save.json"))
            {
                savedFiles = JsonSerializer.Deserialize<Dictionary<string, FileData>>(File.ReadAllText(appPath + "save.json"));
                savedFilesLong = JsonSerializer.Deserialize<Dictionary<string, LongDataSave>>(File.ReadAllText(appPath + "longSave.json"));
            }
            if (!File.Exists(appPath + "fileTitles.json"))
            {
                Console.WriteLine("Generating File Titles");
                GenerateFileTitles();
            }
            else
            {
                savedFileTitles = JsonSerializer.Deserialize<Dictionary<string, SheetData>>(File.ReadAllText(appPath + "fileTitles.json"));
            }
            if (savedFiles.Count != 0)
            {
                Console.WriteLine("Mentés megtalálva. Indítsunk összehasonlítást? (Y/N)");
                //var response = Console.ReadLine().ToLower();
                var response = "y";
                if (response == "y")
                    generateComparison = true;
            }
            Dictionary<string, FileData> fileDatas = GenerateFileData();
            if (fileDatas == null || fileDatas.Count == 0)
            {
                Console.ReadLine();
                return;
            }
            ReportGeneration(fileDatas, generateComparison, savedFiles);

            Console.WriteLine("Saving Data To JSON...");
            var json = JsonSerializer.Serialize(fileDatas);
            File.WriteAllText(appPath + "save.json", json);
            json = JsonSerializer.Serialize(savedFilesLong);
            File.WriteAllText(appPath + "longSave.json", json);
            json = JsonSerializer.Serialize(savedFileTitles);
            File.WriteAllText(appPath + "fileTitles.json", json);
            SaveToSmb("Gerison Check.xlsx", "FGergo\\Gerison Saves\\");

            Console.WriteLine("Save complete!");
            Console.ReadLine();*/
        }

        private static void SaveToSmb(string saveFile, string saveFolder)
        {
            //This function works on black magic
            SMB2Client client = new SMB2Client();
            bool isConnected = client.Connect(IPAddress.Parse("131.0.2.20"), SMBTransportType.DirectTCPTransport);
            if (isConnected)
            {
                NTStatus statuss = client.Login(String.Empty, "brobert", "ma1beSu5");
                if (statuss == NTStatus.STATUS_SUCCESS)
                {
                    Console.WriteLine("woo");
                }
            }

            ISMBFileStore fileStore = client.TreeConnect(@"beszerzes", out NTStatus stat);
            if (fileStore is SMB2FileStore == false)
                return;
            object fileHandle;
            FileStatus fileStatus;
            NTStatus status = fileStore.CreateFile(out fileHandle, out fileStatus, saveFolder + saveFile, AccessMask.GENERIC_WRITE | AccessMask.DELETE | AccessMask.SYNCHRONIZE, SMBLibrary.FileAttributes.Normal, ShareAccess.None, CreateDisposition.FILE_OPEN, CreateOptions.FILE_NON_DIRECTORY_FILE | CreateOptions.FILE_SYNCHRONOUS_IO_ALERT, null);

            //Turns out, there is no owerwrite in this SMB library.
            //So we first delete the old file
            if (status == NTStatus.STATUS_SUCCESS)
            {
                FileDispositionInformation fileDispositionInformation = new FileDispositionInformation();
                fileDispositionInformation.DeletePending = true;
                status = fileStore.SetFileInformation(fileHandle, fileDispositionInformation);
                bool deleteSucceeded = (status == NTStatus.STATUS_SUCCESS);
                status = fileStore.CloseFile(fileHandle);
            }

            //We should be making a backup, "old file", but it's a hassle just to save one file
            //If only I had more time to work on this
            status = fileStore.CreateFile(out fileHandle, out fileStatus, saveFolder + saveFile, AccessMask.GENERIC_WRITE | AccessMask.SYNCHRONIZE, SMBLibrary.FileAttributes.Normal, ShareAccess.None, CreateDisposition.FILE_CREATE, CreateOptions.FILE_NON_DIRECTORY_FILE | CreateOptions.FILE_SYNCHRONOUS_IO_ALERT, null);
            if (status == NTStatus.STATUS_SUCCESS)
            {
                int bitesWritten;
                byte[] bites = File.ReadAllBytes(savePath + saveFile);
                status = fileStore.WriteFile(out bitesWritten, fileHandle, 0, bites);
                if (status != NTStatus.STATUS_SUCCESS)
                {
                    throw new Exception("Failed to write to a file!");
                }
                status = fileStore.CloseFile(fileHandle);
            }
            status = fileStore.Disconnect();
            Console.WriteLine("SMB File Writing Complete");
        }

        private static void ReportGeneration(Dictionary<string, FileData> fileDatas, bool generateComparison, Dictionary<string, FileData> savedFiles)
        {
            Console.WriteLine("Generating Report...");
            if (File.Exists(savePath + "Gerison Check.xlsx"))
            {
                Console.WriteLine("Old file found, renaming it...");
                if (File.Exists(savePath + "Gerison Check_old.xlsx"))
                    File.Delete(savePath + "Gerison Check_old.xlsx");
                File.Move(savePath + "Gerison Check.xlsx", savePath + "Gerison Check_old.xlsx");
            }

            var workbook = new XLWorkbook();
            var newDataSheet = workbook.AddWorksheet("Friss Adatok");

            int newRow = 2;

            newDataSheet.Cell(1, 1).Value = "Fájlnév";
            newDataSheet.Cell(1, 1).Style.Font.Bold = true;
            newDataSheet.Cell(1, 2).Value = "Beszállító neve";
            newDataSheet.Cell(1, 2).Style.Font.Bold = true;
            newDataSheet.Cell(1, 3).Value = "Előző Méret";
            newDataSheet.Cell(1, 3).Style.Font.Bold = true;
            newDataSheet.Cell(1, 4).Value = "Jelenlegi Méret";
            newDataSheet.Cell(1, 4).Style.Font.Bold = true;
            newDataSheet.Cell(1, 5).Value = "Utolsó módosítás";
            newDataSheet.Cell(1, 5).Style.Font.Bold = true;
            newDataSheet.Cell(1, 6).Value = "Új módosítás";
            newDataSheet.Cell(1, 6).Style.Font.Bold = true;
            newDataSheet.Cell(1, 7).Value = "17:00 és 19:45 között?";
            newDataSheet.Cell(1, 7).Style.Font.Bold = true;

            FileData hipolSpecial = new FileData();
            foreach (var data in fileDatas.Values)
            {
                if (!savedFiles.ContainsKey(data.fileName))
                {
                    Console.WriteLine("Fájl mentés nem található: {0}", data.fileName);
                    continue;
                }
                if (data.fileName.Contains("hipol"))
                {
                    //"Nope" is the default empty value
                    if (hipolSpecial.fileName == "Nope")
                    { hipolSpecial = data; continue; }

                    if (DateTime.Compare(data.creationDate, hipolSpecial.creationDate) > 0)
                        hipolSpecial = data; continue;
                }

                if (data.lastModifiedDate.Date < new DateTime(2024, 10, 1, 00, 00, 0)) continue;

                newDataSheet.Cell(newRow, 1).Value = data.fileName;
                if (savedFileTitles.ContainsKey(data.fileName))
                    newDataSheet.Cell(newRow, 2).Value = savedFileTitles[data.fileName].compName;
                newDataSheet.Cell(newRow, 3).Value = (savedFiles[data.fileName].fileSize / 1024 / 1024).ToString("0.00") + " MB";
                newDataSheet.Cell(newRow, 4).Value = (data.fileSize / 1024 / 1024).ToString("0.00") + " MB";
                newDataSheet.Cell(newRow, 5).Value = savedFilesLong[data.fileName].updates.Last().ToShortDateString() + " " + savedFilesLong[data.fileName].updates.Last().ToShortTimeString();
                newDataSheet.Cell(newRow, 6).Value = data.lastModifiedDate.ToShortDateString() + " " + data.lastModifiedDate.ToShortTimeString();
                if (IsDateOld(data.lastModifiedDate))
                    newDataSheet.Range(newDataSheet.Cell(newRow, 1), newDataSheet.Cell(newRow, 7)).Style.Fill.BackgroundColor = XLColor.LightPastelPurple;

                float sizeMb = (savedFiles[data.fileName].fileSize - data.fileSize) / 1024 / 1024;
                if (sizeMb < -10 || sizeMb > 10)
                    newDataSheet.Cell(newRow, 4).Style.Fill.BackgroundColor = XLColor.Red;
                if (data.fileSize == 0)
                    newDataSheet.Cell(newRow, 4).Style.Fill.BackgroundColor = XLColor.Red;

                if (
                    TimeSpan.Compare(data.creationDate.TimeOfDay, compareTime.TimeOfDay) == 1 &&
                    TimeSpan.Compare(data.creationDate.TimeOfDay, compareTimeTo.TimeOfDay) == -1
                )
                {
                    newDataSheet.Cell(newRow, 7).Value = "IGAZ";
                }
                else
                {
                    newDataSheet.Cell(newRow, 7).Value = "HAMIS";
                }

                //We're getting rid of these values at the end of the program
                //We're using these to sort the dictionary, because normally you can't sort a dictionary
                if (savedFileTitles.ContainsKey(data.fileName))
                    newDataSheet.Cell(newRow, 10).Value = savedFileTitles[data.fileName].important;
                newRow++;
            }

            if (!generateComparison)
            {
                workbook.SaveAs(savePath + "Gerison Check.xlsx");
                Console.WriteLine("Report Generation Complete!");
            }

            foreach (var item in fileDatas)
            {
                //We exclude hipol, because we don't like hipol
                //We're keeping the most up to date tho
                if (item.Key.Contains("hipol")) continue;
                if (savedFilesLong.ContainsKey(item.Key))
                {

                    if (savedFilesLong[item.Key].updates.Contains(item.Value.creationDate))
                        continue;
                    savedFilesLong[item.Key].updates.Add(item.Value.creationDate);
                    if (savedFilesLong[item.Key].updates.Count > 20)
                        savedFilesLong[item.Key].updates.RemoveAt(0);
                }
                else
                {
                    LongDataSave save = new LongDataSave();
                    save.updates = new List<DateTime> { item.Value.creationDate };
                    save.fileName = item.Key;
                    savedFilesLong.Add(item.Key, save);
                }
            }

            //int x = GenerateAverage(workbook);

            //IXLWorksheet compareSheet = GenerateComparisons(fileDatas, savedFiles, workbook);

            var rows = newDataSheet.Range("A2:J10000");
            rows.Sort("10 DESC");
            newDataSheet.Range("J:J").Delete(XLShiftDeletedCells.ShiftCellsLeft);
            newDataSheet.Row(savedFileTitles.Values.Count(x => x.important == true) + 1).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            workbook.SaveAs(savePath + "Gerison Check.xlsx");
            Console.WriteLine("Report Generation Complete!");
        }

        private static bool IsDateOld(DateTime date)
        {
            var treshold = DateTime.Now.AddHours(-30);
            return date < treshold;
        }

        private static IXLWorksheet GenerateComparisons(Dictionary<string, FileData> fileDatas, Dictionary<string, FileData> savedFiles, XLWorkbook workbook)
        {
            var compareSheet = workbook.AddWorksheet("Összehasonlítás");

            List<FileData> removedFiles = new List<FileData>();
            List<FileData[]> modifiedFiles = new List<FileData[]>();
            List<FileData> newData = new List<FileData>();
            Console.WriteLine("Saved Files: {0}", savedFiles.Count);
            foreach (var oldData in savedFiles)
            {
                if (!fileDatas.ContainsKey(oldData.Key))
                {
                    removedFiles.Add(oldData.Value);
                    continue;
                }
                if (
                    oldData.Value.fileSize != fileDatas[oldData.Key].fileSize ||
                    oldData.Value.creationDate != fileDatas[oldData.Key].creationDate ||
                    oldData.Value.lastModifiedDate != fileDatas[oldData.Key].creationDate
                ) modifiedFiles.Add(new FileData[] { oldData.Value, fileDatas[oldData.Key] });
            }
            foreach (var fileData in fileDatas)
                if (!savedFiles.ContainsKey(fileData.Key))
                    newData.Add(fileData.Value);

            compareSheet.Cell(1, 1).Value = "Fájl";
            compareSheet.Cell(1, 2).Value = "Beszállító neve";
            compareSheet.Cell(1, 3).Value = "Régi dátum";
            compareSheet.Cell(1, 4).Value = "Új dátum";
            compareSheet.Cell(1, 5).Value = "Fájlméret Különbség (MB)";
            compareSheet.Cell(1, 6).Value = "Törölt fájlok";
            compareSheet.Cell(1, 7).Value = "Új fájlok";

            int x = 2;
            foreach (FileData[] data in modifiedFiles)
            {
                compareSheet.Cell(x, 1).Value = data[0].fileName;
                if (savedFileTitles.ContainsKey(data[0].fileName))
                    compareSheet.Cell(x, 2).Value = savedFileTitles[data[0].fileName].compName;
                compareSheet.Cell(x, 3).Value = data[0].creationDate.Date;
                compareSheet.Cell(x, 4).Value = data[1].creationDate.Date;
                compareSheet.Cell(x, 5).Value = (data[0].fileSize - data[1].fileSize) / 1024 / 1024;
                if (savedFileTitles.ContainsKey(data[0].fileName))
                    compareSheet.Cell(x, 10).Value = savedFileTitles[data[0].fileName].important;
                else
                    compareSheet.Cell(x, 10).Value = false;
                x++;
            }

            x = 2;
            foreach (FileData data in removedFiles)
            {
                compareSheet.Cell(x, 6).Value = data.fileName;
                x++;
            }

            x = 2;
            foreach (FileData data in newData)
            {
                compareSheet.Cell(x, 7).Value = data.fileName;
                x++;
            }

            return compareSheet;
        }

        private static int GenerateAverage(XLWorkbook workbook)
        {
            var averageSheet = workbook.AddWorksheet("Átlagolás");

            averageSheet.Cell(1, 1).Value = "Fájlnév";
            averageSheet.Cell(1, 2).Value = "Beszállító Neve";
            averageSheet.Cell(1, 3).Value = "Átlagos frissítés";
            averageSheet.Cell(1, 4).Value = "Dátumok";
            int x = 2;
            foreach (var item in savedFilesLong)
            {
                long average = Convert.ToInt64(item.Value.updates.Average(date => date.TimeOfDay.Ticks));
                TimeSpan averageTime = new TimeSpan(average);

                averageSheet.Cell(x, 1).Value = item.Key;
                if (savedFileTitles.ContainsKey(item.Key))
                    averageSheet.Cell(x, 2).Value = savedFileTitles[item.Value.fileName].compName;
                averageSheet.Cell(x, 3).Value = averageTime;

                int y = 4;
                foreach (var date in item.Value.updates)
                {
                    averageSheet.Cell(x, y).Value = date.ToShortTimeString();
                    y++;
                }
                x++;
            }

            return x;
        }

        private static Dictionary<string, FileData> GenerateFileData()
        {
            Console.WriteLine("Fetching data...");
            List<string> allFiles = Directory.GetFiles(filePath).ToList();
            List<string> validFiles = new List<string>();
            foreach (string file in allFiles)
            {
                string ext = Path.GetExtension(file);
                string name = Path.GetFileName(file);
                if (ext == ".csv" || ext == ".txt" || ext == ".xlsx")
                    if (!name.Contains("zip"))
                        validFiles.Add(file);
            }

            List<FileData> fileDatas = new List<FileData>();
            foreach (string file in validFiles)
            {
                FileData data = new FileData(
                    Path.GetFileName(file),
                    new System.IO.FileInfo(file).Length,
                    File.GetCreationTime(file),
                    File.GetLastWriteTime(file)
                );
                fileDatas.Add(data);
            }

            Dictionary<string, FileData> returnData = new Dictionary<string, FileData>();
            foreach (FileData file in fileDatas)
                returnData.Add(file.fileName, file);

            if (returnData.Count > 0)
                Console.WriteLine("Data Fetch Complete!");
            else
            { Console.WriteLine("Data Fetch Compromised!"); return null; }

            return returnData;
        }

        private static void GenerateFileTitles()
        {
            StreamReader sr = new StreamReader(appPath + "fajl_beszall.csv");
            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine();
                line = line.Replace("\"", "");
                string[] values = line.Split(';');

                if (values[5] == "FALSE") continue;

                SheetData sheetData = new SheetData(
                    values[0],
                    values[1],
                    values[6]
                );
                savedFileTitles.Add(values[0], sheetData);
            }
        }
    }

    struct SheetData
    {
        public string fileName { get; set; }
        public string compName { get; set; }
        public bool important { get; set; }
        public SheetData(string fileName, string compName, string important)
        {
            this.fileName = fileName;
            this.compName = compName;
            this.important = (important == "TRUE") ? true : false;
        }
    }

    struct FileData
    {
        public string fileName { get; set; }
        public float fileSize { get; set; }
        public DateTime creationDate { get; set; }
        public DateTime lastModifiedDate { get; set; }

        public FileData(string fileName, float fileSize, DateTime creationDate, DateTime lastModifiedDate)
        {
            this.fileName = fileName;
            this.fileSize = fileSize;
            this.creationDate = creationDate;
            this.lastModifiedDate = lastModifiedDate;

            if (fileName == null)
                this.fileName = "Nope";
        }
    }

    struct LongDataSave
    {
        public string fileName { get; set; }
        public List<DateTime> updates { get; set; }
    }
}
