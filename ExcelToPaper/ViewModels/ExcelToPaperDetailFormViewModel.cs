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
using MaterialDesignThemes.Wpf;
using System.Threading;

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
        private CancellationTokenSource mCancelTokenSource = null;
        private PrintSettings mPrintSettings = new PrintSettings();
        //Export to a single folder
        public bool ExportToSingleFolder { get; set; } = false;
        //Attach workbook name before work sheet name
        public bool AttachWorkbookNameBeforeWorksheet { get; set; } = false;
        //Cancel button visibility
        public bool IsCanelButtonVisiable { get; set; } = false;

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
        //The single folder path that is used to be exported.
        public string SingleFolderPath { get; set; }

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
            get
            {
                return mSelectAllSheet;
            }
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

        public ICommand ExportSettingCommand => new DelegateCommand(async (O) => {
            var vm = new ExcelToPaperExportSelectionViewModel();
            //Init old value.
            vm.ExportToSingleFolder = mPrintSettings.ExportToSingleFolder;
            vm.SingleFolderPath = mPrintSettings.SingleFolderPath;
            vm.AttachWorkbookNameBeforeWorksheet = mPrintSettings.AttachWorkbookNameBeforeWorksheet;
            vm.PrintToPaper = mPrintSettings.PrintToPaper;
            //Show dialog.
            var result = await DialogHost.Show(vm.View, "Root");
            var dialogResult = (bool)result;
            if (!dialogResult) return;

            mPrintSettings.ExportToSingleFolder = vm.ExportToSingleFolder;
            mPrintSettings.AttachWorkbookNameBeforeWorksheet = vm.AttachWorkbookNameBeforeWorksheet;
            mPrintSettings.SingleFolderPath = vm.SingleFolderPath;
            mPrintSettings.PrintToPaper = vm.PrintToPaper;
        });

        public ICommand GetDetailInformationCommand => new DelegateCommand(async (o) =>
        {
            //Show dialog
            var vm = new ExcelToPaperGetDetailSelectionViewModel();
            var result = await DialogHost.Show(vm.View, "MiddleRight");
            var dialogResult = (bool)result;
            if (!dialogResult) return;

            //Show progress bar
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
                if (!vm.GetAllFileDetail)
                {
                    if (SelectedFilePath.IsNullOrEmpty())
                        return;
                    await CommonMethods.GetWorksheetPageCount(mExcel, SelectedFilePath, SheetInfos, UpdateProgressMessage);
                }
                else
                    foreach (var kvp in ExcelInfos)
                        await CommonMethods.GetWorksheetPageCount(mExcel, kvp.Key, kvp.Value, UpdateProgressMessage);
            }

            GetDetailProgressValue = 100;
            NotifyPropertyChanged(nameof(GetDetailProgressValue));
        });

        public ICommand ExcelToPaperCancelCommand => new DelegateCommand((o) => {
            if (mCancelTokenSource == null)
                return;
            mCancelTokenSource.Cancel();
        });

        public ICommand OkCommand => new DelegateCommand(async (o) =>
        {
            //Check the cancel token source.
            if (mCancelTokenSource != null)
                return;

            //Init cancel token.
            mCancelTokenSource = new CancellationTokenSource();
            //Make the cancel button visible.
            IsCanelButtonVisiable = true;
            NotifyPropertyChanged(nameof(IsCanelButtonVisiable));

            //Set progress to 0 to make the progress bar show
            ProgressValue = 0;
            NotifyPropertyChanged(nameof(ProgressValue));
                        
            //Wait until export finish.
            await ExcelToPaperMethods.PrintToPaper(
                SelectedPrinter,
                mCancelTokenSource.Token,
                ExcelInfos, 
                mPrintSettings,
                UpdateProgressMessage);

            //Set finish message
            ProgressMessage = "完成";
            NotifyPropertyChanged(nameof(ProgressMessage));

            //Dispose cancel token source
            mCancelTokenSource.Dispose();
            mCancelTokenSource = null;

            //Make the cancel button invisible.
            IsCanelButtonVisiable = false;
            NotifyPropertyChanged(nameof(IsCanelButtonVisiable));

            //Wait 5s and clear finish message, hide progress bar
            await Task.Delay(5000);
            ProgressMessage = "";
            ProgressValue = 100;
            NotifyPropertyChanged(nameof(ProgressMessage));
            NotifyPropertyChanged(nameof(ProgressValue));
        });

        public ICommand CancelCommand => new DelegateCommand((o) => {
            //Can not close form when processing a task.
            if (mCancelTokenSource != null)
                return;
            else
                View.Close();
        });

        private void UpdateProgressMessage(string message)
        {
            ProgressMessage = message;
            NotifyPropertyChanged(nameof(ProgressMessage));
        }
    }
}