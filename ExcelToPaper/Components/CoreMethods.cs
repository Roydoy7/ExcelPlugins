using CommonTools;
using ExcelToPaper.DataModels;
using ExcelToPaper.Parameters;
using Microsoft.Office.Interop.Excel;
using OpenXmlExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToPaper.Components
{
    public static class CoreMethods
    {
        #region Worksheet names
        //Get worksheet names of a workbook
        public static IEnumerable<string> GetWorksheetNames(this WorkbookInfo workbookInfo)
        {
            return GetWorksheetNames(workbookInfo.FilePath);
        }

        //Get worksheet names of a workbook
        private static IEnumerable<string> GetWorksheetNames(string filePath)
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
        #endregion

        #region Excel path
        //Get all the workbooks of a folder
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
        #endregion

        #region Worksheet size count preview

        //Calc worksheet page size and count
        internal static async Task<bool> GetWorksheetPageCountAndSize(this WorkbookInfo workbookInfo, Application excel, CancellationToken cancellationToken, Action<string> updateStatus = null)
        {
            return await OperateWorkbook(excel, workbookInfo, (wb) =>
            {
                foreach (Worksheet ws in wb.Worksheets)
                {
                    if (workbookInfo.WorksheetInfos.Any(x => x.SheetName == ws.Name))
                    {
                        var target = workbookInfo.WorksheetInfos.First(x => x.SheetName == ws.Name);
                        target.Count = ws.PageSetup.Pages.Count;
                        target.PaperSize = ws.PageSetup.PaperSize;
                        target.Orientation = ws.PageSetup.Orientation;
                        target.NotifyPropertyChanged(nameof(target.Count));
                        target.NotifyPropertyChanged(nameof(target.PaperSize));
                    }
                    cancellationToken.ThrowIfCancellationRequested();
                }

                wb?.Close(SaveChanges: false);
                return true;
            }, updateStatus);
        }

        //Get previews of worksheets
        public static async Task<bool> GetWorkSheetPreview(this WorkbookInfo workbookInfo, Application excel, CancellationToken cancellationToken, Action<string> updateStatus = null)
        {
            return await OperateWorkbook(excel, workbookInfo, (wb) =>
            {
                foreach (Worksheet ws in wb.Worksheets)
                {
                    //Ignore hidden worksheet
                    if (ws.Visible != XlSheetVisibility.xlSheetVisible)
                        continue;
                    if (workbookInfo.WorksheetInfos.Any(x => x.SheetName == ws.Name))
                    {
                        //Find target worksheet info
                        var target = workbookInfo.WorksheetInfos.First(x => x.SheetName == ws.Name);
                        target.PreviewsRaw.Clear();

                        //If page count is zero, ignore
                        if (target.Count == 0)
                        {
                            if (ws.PageSetup.Pages.Count == 0)
                                continue;
                        }

                        //Export as pdf
                        var pdfFilePath = CreatePdfPath();
                        ws.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfFilePath);

                        //Add bitmap into preview raw data
                        foreach (var bmp in PdfToBitmapMethods.ToBitmaps(pdfFilePath))
                            target.PreviewsRaw.Add(bmp);

                        //Delete temp pdf file
                        if (File.Exists(pdfFilePath))
                            File.Delete(pdfFilePath);
                    }
                    cancellationToken.ThrowIfCancellationRequested();
                }
                return true;
            }, updateStatus);
        }

        //HOF shell method
        //Used by GetWorksheetPageCountAndSize, GetWorkSheetPreview
        private static async Task<bool> OperateWorkbook(
            Application excel,
            WorkbookInfo workbookInfo,
            Func<Workbook, bool> func,
            Action<string> updateStatus = null)
        {
            var result = false;
            if (excel == null)
                return result;

            await Task.Run(async () =>
            {
                Workbook wb = null;
                try
                {
                    wb = excel.Workbooks.Open(Filename: workbookInfo.FilePath, ReadOnly: true);
                    var retry = 0;

                START:
                    try
                    {
                        if (wb == null) return;
                        try
                        {
                            //May throw operation cancel exception
                            result = func(wb);
                        }
                        catch (COMException e)
                        {
                            //RPC_E_SERVERCALL_RETRYLATER
                            if (e.ErrorCode == -2147417846)
                            {
                                var rand = new Random();
                                //Wait for random time and retry
                                await Task.Delay(rand.Next(100));
                                retry++;
                                //Try no more than 5 times
                                if (retry < 5)
                                    goto START;
                            }
                        }
                    }
                    catch (OperationCanceledException e)
                    {
                        wb?.Close(SaveChanges: false);
                        updateStatus?.Invoke(e.Message);
                        await Task.Delay(2000);
                        updateStatus?.Invoke("");
                    }
                }
                catch (Exception e)
                {
                    wb?.Close(SaveChanges: false);
                    updateStatus?.Invoke(e.Message);
                    //Wait for 5s and clear message.
                    await Task.Delay(5000);
                    updateStatus?.Invoke("");
                }
            });

            return result;
        }
        #endregion

        #region Pdf path
        //Create a pdf path with a random name
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
        #endregion
    }
}