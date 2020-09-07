using CommonTools;
using CommonTools.FileBrowsers;
using CommonWPFTools;
using ExcelToPaper.Components;
using ExcelToPaper.DataModels;
using ExcelToPaper.Views;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelToPaper.ViewModels
{
    internal class ExcelToPaperDetailFormViewModel : ViewModelBase<ExcelToPaperDetailForm>
    {
        public ExcelToPaperDetailFormViewModel()
        {
            //Set default printer
            var printSettings = new PrinterSettings();
            SelectedPrinter = printSettings.PrinterName;
        }

        ~ExcelToPaperDetailFormViewModel()
        {
            if (mExcel != null)
            {
                mExcel.Quit();
                mExcel = null;
            }
        }

        private static Application mExcel = null;
        private bool mSelectAllSheet;

        //A dictionary that contain excel info,
        //the key is path of the excel file,
        //the value is a list of sheet information,
        //the element of the sheet information list contains sheet name and checked boolean.
        public Dictionary<string, List<SheetInfo>> ExcelInfos { get; set; } = new Dictionary<string, List<SheetInfo>>();

        public ObservableCollection<string> ExcelFilePaths { get; set; } = new ObservableCollection<string>();
        public int ProgressValue { get; set; } = 100;
        public int GetDetailProgressValue { get; set; } = 100;
        public uint StartPage { get; set; }
        public uint EndPage { get; set; }
        public string Keyword { get; set; } = "";
        public string SelectedFilePath { get; set; }
        public string SelectedPrinter { get; set; }
        public string ProgressMessage { get; set; }

        public string SelectedSheetMessage
        {
            get
            {
                var cnt = 0;
                foreach (var kvp in ExcelInfos)
                {
                    foreach (var sheetInfo in kvp.Value)
                        if (sheetInfo.IsSheetChecked)
                            cnt++;
                }
                return $"選択されたシート数: {cnt}";
            }
        }

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

                NotifyPropertyChanged(nameof(SelectedSheetMessage));
            }
        }

        public IEnumerable<string> InstalledPrinters
        {
            get
            {
                var printerList = new List<string>();
                foreach (var printer in PrinterSettings.InstalledPrinters)
                    printerList.Add(printer as string);
                return printerList;
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

        //Open export folder
        public ICommand OpenExportFolderCommand => new DelegateCommand(async (o) =>
        {
            if (SelectedFilePath == null)
                return;
            var folderPath = Path.GetDirectoryName(SelectedFilePath);
            var fileName = Path.GetFileNameWithoutExtension(SelectedFilePath);
            var exportFolderPath = folderPath + "\\" + fileName;
            if (Directory.Exists(exportFolderPath))
                Process.Start("explorer.exe", exportFolderPath);
            else
            {
                UpdateProgressMessage("フォルダが存在しません.");
                await Task.Delay(5000);
                UpdateProgressMessage("");
            }
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

        public ICommand SingleSheetCheckedCommand => new DelegateCommand((o) =>
        {
            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand SelectByKeywordCommand => new DelegateCommand((o) =>
        {
            if (Keyword.IsNullOrEmpty())
                return;
            foreach (var kvp in ExcelInfos)
                foreach (var sheetInfo in kvp.Value)
                {
                    if (sheetInfo.SheetName.ToUpper().Contains(Keyword.ToUpper()))
                    {
                        sheetInfo.IsSheetChecked = true;
                        sheetInfo.NotifyPropertyChanged(nameof(sheetInfo.IsSheetChecked));
                    }
                }

            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand UnSelectByKeywordCommand => new DelegateCommand((o) =>
        {
            if (Keyword.IsNullOrEmpty())
                return;
            foreach (var kvp in ExcelInfos)
                foreach (var sheetInfo in kvp.Value)
                {
                    if (sheetInfo.SheetName.ToUpper().Contains(Keyword.ToUpper()))
                    {
                        sheetInfo.IsSheetChecked = false;
                        sheetInfo.NotifyPropertyChanged(nameof(sheetInfo.IsSheetChecked));
                    }
                }

            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand SelectAllCommand => new DelegateCommand((o) =>
        {
            foreach (var kvp in ExcelInfos)
                foreach (var sheetInfo in kvp.Value)
                {
                    sheetInfo.IsSheetChecked = true;
                    sheetInfo.NotifyPropertyChanged(nameof(sheetInfo.IsSheetChecked));
                }

            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand UnSelectAllCommand => new DelegateCommand((o) =>
        {
            foreach (var kvp in ExcelInfos)
                foreach (var sheetInfo in kvp.Value)
                {
                    sheetInfo.IsSheetChecked = false;
                    sheetInfo.NotifyPropertyChanged(nameof(sheetInfo.IsSheetChecked));
                }

            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand GetDetailInformationCommand => new DelegateCommand(async (o) =>
        {
            if (SelectedFilePath.IsNullOrEmpty())
                return;

            GetDetailProgressValue = 0;
            NotifyPropertyChanged(nameof(GetDetailProgressValue));

            if (mExcel == null)
                await Task.Run(() =>
                {
                    mExcel = new Application();
                    mExcel.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                });

            if (mExcel != null)
            {
                await CommonMethods.GetWorksheetPageCount(mExcel, SelectedFilePath, SheetInfos, UpdateProgressMessage);
            }

            GetDetailProgressValue = 100;
            NotifyPropertyChanged(nameof(GetDetailProgressValue));
        });

        public ICommand OkCommand => new DelegateCommand(async (o) =>
        {
            //Set progress to 0 to make the progress bar show
            ProgressValue = 0;
            NotifyPropertyChanged(nameof(ProgressValue));

            //Wait until export finish.
            await ExcelToPaperMethods.PrintToPaper(SelectedPrinter, ExcelInfos, UpdateProgressMessage);

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

        private void UpdateProgressMessage(string message)
        {
            ProgressMessage = message;
            NotifyPropertyChanged(nameof(ProgressMessage));
        }
    }
}