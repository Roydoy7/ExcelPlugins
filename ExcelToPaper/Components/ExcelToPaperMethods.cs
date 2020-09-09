using CommonTools;
using ExcelToPaper.DataModels;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System;

namespace ExcelToPaper.Components
{
    class ExcelToPaperMethods
    {
        public static async Task PrintToPaper(
            string printer, //Printer name
            CancellationToken cancelToken,//Cancel token
            Dictionary<string, List<SheetInfo>> excelInfos, //Workbook and worksheet info.
            PrintSettings printSettings,
            Action<string> updateStatus = null//Invoke to update status
            )
        {
            //Valid check
            if (printSettings.ExportToSingleFolder)
                if (printSettings.SingleFolderPath.IsNullOrEmpty())
                    return;
            else
                {
                    //Remove \\ at the end
                    if (printSettings.SingleFolderPath.EndsWith("\\"))
                        printSettings.SingleFolderPath = printSettings.SingleFolderPath.TrimEnd('\\');
                }

            await Task.Run(async () =>
            {
                //Start excel
                var excel = new Application();
                //Print out
                foreach (var kvp in excelInfos)
                {
                    updateStatus?.Invoke($"処理中 {Path.GetFileName(kvp.Key)}...");
                    PrintToPaper(
                        excel, 
                        printer,
                        cancelToken,
                        kvp.Key, 
                        kvp.Value,
                        printSettings
                        );
                    if(cancelToken.IsCancellationRequested)
                    {
                        updateStatus?.Invoke($"中止中...");
                        await Task.Delay(500);
                        break;
                    }
                }
                //Quit
                excel.Quit();
            });
        }

        private static void PrintToPaper(
            Application excel, 
            string printer,
            CancellationToken cancelToken,//Cancel token
            string filePath, 
            List<SheetInfo> sheetInfos,
            PrintSettings printSettings,
            string singleFolderPath = "" //The single folder path to be exported.
            )
        {
            //Check if there is any worksheet to print
            if (!sheetInfos.Any(x => x.IsSheetChecked))
                return;

            var exportFolderPath = "";
            var pdfNamePrefix = "";

            if (!printSettings.ExportToSingleFolder)
            {
                //Create a folder with the same name as the excel
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var folderPath = Path.GetDirectoryName(filePath);
                exportFolderPath = folderPath + "\\" + fileName;
                Directory.CreateDirectory(exportFolderPath);
            }
            else
            {
                //User single folder path as the export path
                exportFolderPath = printSettings.SingleFolderPath;
                if(printSettings.AttachWorkbookNameBeforeWorksheet)
                {
                    var excelName = Path.GetFileNameWithoutExtension(filePath);
                    pdfNamePrefix = excelName + "_";
                }
            }

            //Try to export worksheet as pdf
            try
            {
                var wb = excel.Workbooks.Open(Filename: filePath, ReadOnly: true);
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
                        //Print to pdf
                        if(!printSettings.PrintToPaper)
                            ws.PrintOut(
                                From:fromPage,
                                To:toPage,
                                ActivePrinter: printer,
                                PrToFileName: exportFolderPath + "\\" + pdfNamePrefix + ws.Name + ".pdf"
                                );
                        //Print to paper
                        else
                            ws.PrintOut(
                               From: fromPage,
                               To: toPage,
                               ActivePrinter: printer
                               );
                    }
                    if (cancelToken.IsCancellationRequested)
                        break;
                }
                wb.Close(SaveChanges:false);
            }
            catch { }
        }

        private static void Test()
        {
            PrintDocument pd = new PrintDocument();
            PrinterSettings ps = new PrinterSettings();
        }
    }
}
