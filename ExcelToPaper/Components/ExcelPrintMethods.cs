using CommonTools;
using ExcelToPaper.DataModels;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToPaper.Components
{
    public class ExcelPrintMethods
    {
        public static async Task<IEnumerable<PrintResult>> PrintToPaper(
            string printer, //Printer name
            CancellationToken cancelToken,//Cancel token
            IEnumerable<WorkbookInfo> workbookInfos,
            PrintSettings printSettings,
            Action<string> updateStatus = null//Invoke to update status
            )
        {
            var printResults = new BlockingCollection<PrintResult>();

            //Valid check
            if (printSettings.ExportToSingleFolder)
                if (printSettings.SingleFolderPath.IsNullOrEmpty())
                    return printResults;
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
                foreach (var workbookInfo in workbookInfos)
                {
                    //Ignore if all the sheets are unchecked
                    if (!workbookInfo.WorksheetInfos.Any(x => x.IsWorksheetChecked))
                        continue;

                    updateStatus?.Invoke($"処理中 {workbookInfo.FileName}...");
                    var result =
                    PrintToPaper(
                        excel,
                        printer,
                        cancelToken,
                        workbookInfo.FilePath,
                        workbookInfo.WorksheetInfos,
                        printSettings
                        );
                    printResults.Add(result);
                    if (cancelToken.IsCancellationRequested)
                    {
                        updateStatus?.Invoke($"中止中...");
                        await Task.Delay(500);
                        break;
                    }
                }
                //Quit
                excel.Quit();
            });

            return printResults;
        }

        private static PrintResult PrintToPaper(
            Application excel,
            string printer,
            CancellationToken cancelToken,//Cancel token
            string filePath,
            IEnumerable<WorksheetInfo> sheetInfos,
            PrintSettings printSettings
            )
        {
            var result = new PrintResult();
            var exportFolderPath = "";
            var pdfNamePrefix = "";

            //Print to separate folders
            if (!printSettings.ExportToSingleFolder)
            {
                //Create a folder with the same name as the excel
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var folderPath = Path.GetDirectoryName(filePath);
                exportFolderPath = folderPath + "\\" + fileName;
                Directory.CreateDirectory(exportFolderPath);
            }
            //Print to a single folder
            else
            {
                //User single folder path as the export path
                exportFolderPath = printSettings.SingleFolderPath;
                if (printSettings.AttachWorkbookNameBeforeWorksheet)
                {
                    var excelName = Path.GetFileNameWithoutExtension(filePath);
                    pdfNamePrefix = excelName + "_";
                }
            }

            result.OutputFolderPath = exportFolderPath;
            result.WorkbookPath = filePath;

            //Try to export worksheet as pdf
            try
            {
                //Open workbook
                var wb = excel.Workbooks.Open(Filename: filePath, ReadOnly: true);
                //A dictionary to save worksheet data from workbook
                var worksheetDict = new Dictionary<string, Worksheet>();
                foreach (Worksheet ws in wb.Worksheets)
                    worksheetDict.Add(ws.Name, ws);

                foreach(var worksheetInfo in sheetInfos.Where(x=>x.IsWorksheetChecked))
                {
                    //Check if checked worksheet's name is contained in this workbook
                    if (!worksheetDict.ContainsKey(worksheetInfo.SheetName))
                        continue;

                    //Get worksheet
                    var ws = worksheetDict[worksheetInfo.SheetName];

                    //Get start and end information
                    uint fromPage = 1;
                    uint toPage = (uint)ws.PageSetup.Pages.Count;
                    if (worksheetInfo.StartPage > 0)
                        fromPage = worksheetInfo.StartPage;
                    if (worksheetInfo.EndPage > 0)
                        toPage = worksheetInfo.EndPage;

                    //Print to pdf
                    if (!printSettings.PrintToPaper)
                    {
                        var pdfFilePath = exportFolderPath + "\\" + pdfNamePrefix + ws.Name + ".pdf";
                        ws.PrintOut(
                            From: fromPage,
                            To: toPage,
                            ActivePrinter: printer,
                            PrToFileName: pdfFilePath
                            );
                        result.PrintedPdfPaths.Add(pdfFilePath);
                    }
                    //Print to paper
                    else
                        ws.PrintOut(
                           From: fromPage,
                           To: toPage,
                           ActivePrinter: printer
                           );

                    if (cancelToken.IsCancellationRequested)
                        break;
                }

                //foreach (Worksheet ws in wb.Worksheets)
                //{
                //    if (sheetInfos.Any(x => x.SheetName == ws.Name && x.IsSheetChecked))
                //    {
                //        var target = sheetInfos.First(x => x.SheetName == ws.Name && x.IsSheetChecked);
                //        uint fromPage = 1;
                //        uint toPage = (uint)ws.PageSetup.Pages.Count;
                //        if (target.StartPage > 0)
                //            fromPage = target.StartPage;
                //        if (target.EndPage > 0)
                //            toPage = target.EndPage;
                //        //Print to pdf
                //        if (!printSettings.PrintToPaper)
                //        {
                //            var pdfFilePath = exportFolderPath + "\\" + pdfNamePrefix + ws.Name + ".pdf";  
                //            ws.PrintOut(
                //                From: fromPage,
                //                To: toPage,
                //                ActivePrinter: printer,
                //                PrToFileName: pdfFilePath
                //                );
                //            result.PrintedPdfPaths.Add(pdfFilePath);
                //        }
                //        //Print to paper
                //        else
                //            ws.PrintOut(
                //               From: fromPage,
                //               To: toPage,
                //               ActivePrinter: printer
                //               );
                //    }
                //    if (cancelToken.IsCancellationRequested)
                //        break;
                //}

                //Close workbook
                wb.Close(SaveChanges: false);
            }
            catch { }

            return result;
        }
    }
}