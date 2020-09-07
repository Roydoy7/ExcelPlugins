using CommonTools;
using ExcelToPdf.Components;
using ExcelToPdf.DataModels;
using ExcelToPdf.ViewModels;
using OpenXmlExcel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToPdf.Commands
{
    public static partial class RibbonCommands
    {
        //Glue method to export to pdf
        public static void BatchToPdf()
        {
            var folderPath = ShowExcelToPdfForm();
            if (folderPath.IsNullOrEmpty()) return;

            var filePaths = CommonMethods.GetExcelPath(folderPath);
            ShowExcelToPdfDetailForm(filePaths);
        }

        private static string ShowExcelToPdfForm()
        {
            var vm = new ExcelToPdfFormViewModel();
            vm.View.ShowDialog();
            if (vm.View.DialogResult.Value)
                return vm.ExcelFolderPath;
            return "";
        }

        private static void ShowExcelToPdfDetailForm(IEnumerable<string> filePaths)
        {
            var vm = new ExcelToPdfDetailFormViewModel();

            //Get worksheet names
            foreach (var filePath in filePaths)
            {
                vm.ExcelFilePaths.Add(filePath);
                vm.ExcelInfos.Add(filePath, CommonMethods.GetWorksheetNames(filePath).Select(x => new SheetInfo { SheetName = x, IsSheetChecked = false }).ToList());
            }

            vm.View.ShowDialog();
        }
    }
}
