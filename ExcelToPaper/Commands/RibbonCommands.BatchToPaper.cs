using CommonTools;
using ExcelToPaper.Components;
using ExcelToPaper.DataModels;
using ExcelToPaper.ViewModels;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToPaper.Commands
{
    public static class RibbonCommands
    {
        //Glue method
        public static void BatchPrint()
        {
            //Show dialog
            var vm = new ExcelToPaperFormViewModel();
            vm.View.ShowDialog();

            //var folderPath = ShowExcelToPaperForm();
            //if (folderPath.IsNullOrEmpty()) return;

            //var filePaths = CommonMethods.GetExcelPath(folderPath);
            //ShowExcelToPaperDetailForm(filePaths);
        }

        public static string ShowExcelToPaperForm()
        {
            var vm = new PathFormViewModel();
            vm.View.ShowDialog();
            if (vm.View.DialogResult.Value)
                return vm.ExcelFolderPath;
            return "";
        }

        //private static void ShowExcelToPaperDetailForm(IEnumerable<string> filePaths)
        //{
        //    var vm = new ExcelToPaperDetailFormViewModel();

        //    //Get worksheet names from excel workbooks parallelly.
        //    var excelFilePaths = new BlockingCollection<string>();
        //    var excelInfos = new ConcurrentDictionary<string, List<WorksheetInfo>>();
        //    Parallel.ForEach(filePaths, filePath => 
        //    {            
        //        excelFilePaths.Add(filePath);
        //        excelInfos.TryAdd(filePath, CommonMethods.GetWorksheetNames(filePath).Select(x => new WorksheetInfo { SheetName = x, IsSheetChecked = false }).ToList());
        //    });

        //    //Copy to window view model.
        //    foreach (var filePath in excelFilePaths)
        //        vm.ExcelFileInfos.Add(filePath);
        //    foreach(var kvp in excelInfos)
        //        vm.ExcelInfos.Add(kvp.Key, kvp.Value);

        //    //Show dialog
        //    vm.View.ShowDialog();
        //}


    }
}
