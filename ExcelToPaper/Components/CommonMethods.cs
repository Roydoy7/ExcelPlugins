using ExcelToPaper.DataModels;
using ExcelToPaper.Parameters;
using Microsoft.Office.Interop.Excel;
using OpenXmlExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CommonTools;

namespace ExcelToPaper.Components
{
    public static class CommonMethods
    {
        public static IEnumerable<string> GetWorksheetNames(string filePath)
        {
            var sheetNames = new List<string>();
            using (var excel = new OpenXExcel())
            {
                try
                {
                    var wb = excel.OpenWorkbook(filePath, false);
                    foreach (var ws in wb.GetWorksheets())
                    {
                        sheetNames.Add(ws.GetName());
                    }
                    wb.CloseWorkbook();
                }
                catch { }
            }
            return sheetNames;
        }

        public static IEnumerable<string> GetExcelPath(string folderPath)
        {
            if (Directory.Exists(folderPath) == false) yield break;
            foreach (var filePath in Directory.GetFiles(folderPath, "*.xlsx"))
                if (!Path.GetFileName(filePath).StartsWith("~$"))
                    yield return filePath;
            foreach (var filePath in Directory.GetFiles(folderPath, "*.xlsm"))
                if (!Path.GetFileName(filePath).StartsWith("~$"))
                    yield return filePath;
        }

        internal static async Task GetWorksheetPageCount(Application excel, string filePath, IEnumerable<WorksheetInfo> sheetInfos, CancellationToken cancellationToken, Action<string> updateStatus = null)
        {
            if (excel == null)
                return;

            await Task.Run(async () =>
            {
                try
                {
                    var wb = excel.Workbooks.Open(Filename: filePath, ReadOnly: true);
                    try
                    {
                        foreach (Worksheet ws in wb.Worksheets)
                        {
                            if (sheetInfos.Any(x => x.SheetName == ws.Name))
                            {
                                var target = sheetInfos.First(x => x.SheetName == ws.Name);
                                target.Count = ws.PageSetup.Pages.Count;
                                target.PaperSize = ws.PageSetup.PaperSize;
                                target.Orientation = ws.PageSetup.Orientation;
                                target.NotifyPropertyChanged(nameof(target.Count));
                                target.NotifyPropertyChanged(nameof(target.PaperSize));
                            }
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        wb.Close(SaveChanges: false);
                    }
                    catch(OperationCanceledException e)
                    {
                        wb.Close(SaveChanges: false);
                        updateStatus?.Invoke(e.Message);
                        //Wait for 5s and clear message.
                        await Task.Delay(5000);
                        updateStatus?.Invoke("");
                    }
                }
                catch (Exception e)
                {
                    updateStatus?.Invoke(e.Message);
                    //Wait for 5s and clear message.
                    await Task.Delay(5000);
                    updateStatus?.Invoke("");
                }
            });
        }

        public static async Task GetWorkSheetPreview(Application excel, string filePath, IEnumerable<WorksheetInfo> sheetInfos, CancellationToken cancellationToken, Action<string> updateStatus = null)
        {
            if (excel == null)
                return;

            await Task.Run(async () =>
            {
                try
                {
                    var wb = excel.Workbooks.Open(Filename: filePath, ReadOnly: true);
                    try
                    {
                        foreach (Worksheet ws in wb.Worksheets)
                        {
                            if (sheetInfos.Any(x => x.SheetName == ws.Name))
                            {
                                //Export as pdf
                                var pdfFilePath = CreatePdfPath();
                                ws.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfFilePath);

                                //Find target worksheet info
                                var target = sheetInfos.First(x => x.SheetName == ws.Name);
                                target.PreviewsRaw.Clear();
                                //Add bitmap into preview raw data
                                foreach (var bmp in PdfToBitmapMethods.ToBitmaps(pdfFilePath))
                                    target.PreviewsRaw.Add(bmp);

                                //Delete temp pdf file
                                if (File.Exists(pdfFilePath))
                                    File.Delete(pdfFilePath);
                            }
                            cancellationToken.ThrowIfCancellationRequested();
                        }
                        wb.Close(SaveChanges: false);
                    }
                    catch(OperationCanceledException e)
                    {
                        wb.Close(SaveChanges: false);
                        updateStatus?.Invoke(e.Message);
                    }
                }
                catch (Exception e)
                {
                    updateStatus?.Invoke(e.Message);
                    //Wait for 5s and clear message.
                    await Task.Delay(5000);
                    updateStatus?.Invoke("");
                }
            });
        }


        private static string CreatePdfPath()
        {
            //Directory name
            var filePath = Path.Combine(PathEx.GetProgramDataPath(), FolderParameters.CompanryFolderName);
            filePath = Path.Combine(filePath, FolderParameters.AppFloderName);
            filePath = Path.Combine(filePath, FolderParameters.TmpFolderName);

            //Create directory
            Directory.CreateDirectory(filePath);

            //Create file name
            var rand = new Random();
            var fileName = DateTime.Now.ToString("yyyyMMddHHmmss" + rand.Next(1000, 9999));
            filePath = filePath + "\\" + fileName + ".pdf";
            return filePath;
        }

    }
}
