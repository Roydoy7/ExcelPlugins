using ExcelToPaper.DataModels;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToPaper.Components
{
    class ExcelToPaperMethods
    {
        public static async Task PrintToPaper(string printer, Dictionary<string, List<SheetInfo>> excelInfos, Action<string> updateStatus = null)
        {
            await Task.Run(() =>
            {
                var excel = new Application();
                foreach (var kvp in excelInfos)
                {
                    updateStatus?.Invoke($"処理中 {Path.GetFileName(kvp.Key)}...");
                    PrintToPaper(excel, printer, kvp.Key, kvp.Value);
                }
                excel.Quit();
            });
        }

        private static void PrintToPaper(Application excel, string printer, string filePath, List<SheetInfo> sheetInfos)
        {
            //Check if there is any worksheet to print
            if (!sheetInfos.Any(x => x.IsSheetChecked))
                return;

            //Create a folder with the same name as the excel
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var folderPath = Path.GetDirectoryName(filePath);
            var exportFolderPath = folderPath + "\\" + fileName;
            Directory.CreateDirectory(exportFolderPath);

            //Try to export worksheet as pdf
            try
            {
                var wb = excel.Workbooks.Open(filePath);
                foreach (Worksheet ws in wb.Worksheets)
                {
                    if (sheetInfos.Any(x => x.SheetName == ws.Name && x.IsSheetChecked))
                    {
                        var target = sheetInfos.First(x => x.SheetName == ws.Name && x.IsSheetChecked);
                        uint fromPage = 1;
                        uint toPage = (uint)ws.PageSetup.Pages.Count;
                        if (target.StartPage > 0)
                            fromPage = target.StartPage;
                        if (target.EndPage > 0)
                            toPage = target.EndPage;

                        ws.PrintOut(
                            From:fromPage,
                            To:toPage,
                            ActivePrinter: printer,
                            PrToFileName: exportFolderPath + "\\" + ws.Name + ".pdf"
                            );
                    }
                }
                wb.Close();
            }
            catch { }

            //Delete folder if is empty
            //if (Directory.GetFiles(exportFolderPath).Length == 0)
            //    Directory.Delete(exportFolderPath, true);
        }

        private static void Test()
        {
            PrintDocument pd = new PrintDocument();
            PrinterSettings ps = new PrinterSettings();
        }
    }
}
