using ExcelToPdf.DataModels;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToPdf.Components
{
    class ExcelToPdfMethods
    {

        public static async Task PrintToPdf(Dictionary<string, List<SheetInfo>> excelInfos, Action<string> updateStatus = null)
        {
            await Task.Run(() =>
            {
                var excel = new Application();
                foreach (var kvp in excelInfos)
                {
                    updateStatus?.Invoke(Path.GetFileName(kvp.Key));
                    PrintToPdf(excel, kvp.Key, kvp.Value);
                }
                excel.Quit();
            });
        }

        private static void PrintToPdf(Application excel, string filePath, List<SheetInfo> sheetInfos)
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
                        ws.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, exportFolderPath + "\\" + ws.Name + ".pdf");
                }
                wb.Close();
            }
            catch { }
        }

    }
}
