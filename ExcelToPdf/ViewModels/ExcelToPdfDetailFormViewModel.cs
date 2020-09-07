using CommonTools.FileBrowsers;
using CommonWPFTools;
using ExcelToPdf.Components;
using ExcelToPdf.DataModels;
using ExcelToPdf.Views;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelToPdf.ViewModels
{
    internal class ExcelToPdfDetailFormViewModel : ViewModelBase<ExcelToPdfDetailForm>
    {
        private bool mSelectAllSheet;

        //A dictionary that contain excel info,
        //the key is path of the excel file,
        //the value is a list of sheet information,
        //the element of the sheet information list contains sheet name and checked boolean.
        public Dictionary<string, List<SheetInfo>> ExcelInfos { get; set; } = new Dictionary<string, List<SheetInfo>>();
        public ObservableCollection<string> ExcelFilePaths { get; set; } = new ObservableCollection<string>();
        public int ProgressValue { get; set; } = 100;
        public string SelectedFilePath { get; set; }
        public string ProgressMessage { get; set; }

        //Check to select all sheets.
        public bool SelectAllSheet
        {
            get => mSelectAllSheet;
            set
            {
                mSelectAllSheet = value;
                if (SelectedFilePath != null)
                    if (ExcelInfos.ContainsKey(SelectedFilePath))
                        foreach (var sheetInfo in ExcelInfos[SelectedFilePath])
                        {
                            sheetInfo.IsSheetChecked = value;
                            sheetInfo.NotifyPropertyChanged(nameof(sheetInfo.IsSheetChecked));
                        }

            }
        }

        //Sheet info used by the second listview.
        public IEnumerable<SheetInfo> SheetInfos
        {
            get
            {
                if (SelectedFilePath != null)
                    return ExcelInfos[SelectedFilePath];
                return new List<SheetInfo>();
            }
        }

        //On select changed of the first listview, update the datasource of the second listview.
        public ICommand ExcelFilePathSelectionChanged => new DelegateCommand((o) =>
        {
            NotifyPropertyChanged(nameof(SheetInfos));
        });

        public ICommand AddFilePathCommand => new DelegateCommand((o) => 
        {
            foreach (var filePath in FileBrowser.BrowseExcelFile(true))
            {
                if (ExcelInfos.ContainsKey(filePath))
                    continue;
                else
                {
                    ExcelFilePaths.Add(filePath);
                    ExcelInfos.Add(filePath, CommonMethods.GetWorksheetNames(filePath).Select(x => new SheetInfo { SheetName = x, IsSheetChecked = false }).ToList());
                }
            }
        });

        public ICommand DeleteFilePathCommand => new DelegateCommand((o) =>
        {
            if (SelectedFilePath == null)
                return;
            if (ExcelInfos.ContainsKey(SelectedFilePath))
                ExcelInfos.Remove(SelectedFilePath);
            ExcelFilePaths.Remove(SelectedFilePath);
        });

        public ICommand OkCommand => new DelegateCommand(async (o) =>
        {
            //Set progress to 0 to make the progress bar show
            ProgressValue = 0;
            NotifyPropertyChanged(nameof(ProgressValue));

            //Wait until export finish.
            await ExcelToPdfMethods.PrintToPdf(ExcelInfos, UpdateProgressMessage);

            //Set finish message
            ProgressMessage = "完成";
            NotifyPropertyChanged(nameof(ProgressMessage));

            //Wait 5s and clear finish message, hide progress bar
            await Task.Delay(5000);
            ProgressMessage = "";
            ProgressValue = 100;
            NotifyPropertyChanged(nameof(ProgressMessage));
            NotifyPropertyChanged(nameof(ProgressValue));
        });

        private void UpdateProgressMessage(string fileName)
        {
            ProgressMessage = $"処理中 {fileName}...";
            NotifyPropertyChanged(nameof(ProgressMessage));
        }
    }

}