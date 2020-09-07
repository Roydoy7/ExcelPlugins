using ExcelToPdf.DataModels;
using Microsoft.Office.Interop.Excel;
using OpenXmlExcel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToPdf.Components
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
                        sheetNames.Add(ws.GetName(excel));
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
        }

    }
}
