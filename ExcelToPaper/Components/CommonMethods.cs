using ExcelToPaper.DataModels;
using Microsoft.Office.Interop.Excel;
using OpenXmlExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

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
                        sheetNames.Add(ws.GetName(excel));
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
                if(!Path.GetFileName(filePath).StartsWith("~$"))
                    yield return filePath;
            foreach (var filePath in Directory.GetFiles(folderPath, "*.xlsm"))
                if (!Path.GetFileName(filePath).StartsWith("~$"))
                    yield return filePath;
        }

        internal static async Task GetWorksheetPageCount(Application excel, string filePath, IEnumerable<SheetInfo> sheetInfos, Action<string> updateStatus = null)
        {
            if (excel == null)
                return;

            await Task.Run(async () => {
                try
                {
                    var wb = excel.Workbooks.Open(Filename: filePath, ReadOnly: true);
                    foreach (Worksheet ws in wb.Worksheets)
                    {
                        if (sheetInfos.Any(x => x.SheetName == ws.Name))
                        {
                            var target = sheetInfos.First(x => x.SheetName == ws.Name);
                            target.Count = ws.PageSetup.Pages.Count;
                            target.PaperSize = ws.PageSetup.PaperSize;
                            target.NotifyPropertyChanged(nameof(target.Count));
                            target.NotifyPropertyChanged(nameof(target.PaperSize));
                        }
                    }
                    wb.Close(SaveChanges: false);
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
    }
}
