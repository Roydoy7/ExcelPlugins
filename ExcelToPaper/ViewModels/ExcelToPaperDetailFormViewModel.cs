using CommonTools;
using CommonTools.FileBrowsers;
using CommonWPFTools;
using ExcelToPaper.Components;
using ExcelToPaper.DataModels;
using ExcelToPaper.Views;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Excel;
using PropertyChanged;
using SpreadsheetLight;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Threading;
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
            if (Excel != null)
            {
                Excel.Workbooks.Close();
                Excel.Quit();
                Excel = null;
            }
        }

        public static Application Excel { get; private set; }
        public bool ExcelIsBusy { get; set; } = false;
        private bool mSelectAllSheet;
        private CancellationTokenSource mCancelTokenSource = null;
        private PrintSettings mPrintSettings = new PrintSettings();

        //Export to a single folder
        public bool ExportToSingleFolder { get; set; } = false;

        //Attach workbook name before work sheet name
        public bool AttachWorkbookNameBeforeWorksheet { get; set; } = false;

        //Cancel button visibility
        public bool IsCanelButtonVisiable { get; set; } = false;

        public bool ShowPreviewProgressBar { get; set; }
        public bool ShowProgressBar { get; set; }
        public uint StartPage { get; set; }
        public uint EndPage { get; set; }
        public string Keyword { get; set; } = "";
        public string SelectedPrinter { get; set; }
        public string ProgressMessage { get; set; }

        //The single folder path that is used to be exported.
        public string SingleFolderPath { get; set; }

        public WorkbookInfo SelectedWorkbookInfo { get; set; }
        public WorksheetInfo SelectedWorksheetInfo { get; set; }
        public ObservableCollection<WorkbookInfo> WorkbookInfos { get; set; } = new ObservableCollection<WorkbookInfo>();

        public string SelectedSheetMessage
        {
            get
            {
                var cnt = 0;
                foreach (var workbookInfo in WorkbookInfos)
                    foreach (var worksheetInfo in workbookInfo.WorksheetInfos)
                        if (worksheetInfo.IsSheetChecked)
                            cnt++;

                return $"選択されたシート数: {cnt}";
            }
        }

        //Check to select all sheets.
        [AlsoNotifyFor(nameof(SelectedSheetMessage))]
        public bool SelectAllSheet
        {
            get => mSelectAllSheet;
            set
            {
                mSelectAllSheet = value;

                if (SelectedWorkbookInfo != null)
                    foreach (var worksheetInfo in SelectedWorkbookInfo.WorksheetInfos)
                        worksheetInfo.IsSheetChecked = value;
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

        public ICommand AddFromFolderCommand => new DelegateCommand(async (o) =>
        {
            if (ExcelIsBusy)
            {
                UpdateProgressMessage("Excelは忙しいです。");
                await Task.Delay(2000);
                UpdateProgressMessage("");
                return;
            }

            var folderPath = FileBrowser.BrowseFolder();
            if (folderPath.IsNullOrEmpty())
                return;

            //Show progress bar
            ShowProgressBar = true;

            //Get worksheet names from excel workbooks parallelly.
            var filePaths = CommonMethods.GetExcelPath(folderPath);
            var workbookInfos = new BlockingCollection<WorkbookInfo>();
            Parallel.ForEach(filePaths, filePath =>
            {
                var workbookInfo = new WorkbookInfo { FilePath = filePath };
                foreach (var sheetName in CommonMethods.GetWorksheetNames(filePath))
                    workbookInfo.WorksheetInfos.Add(new WorksheetInfo { SheetName = sheetName });
                workbookInfos.Add(workbookInfo);
            });

            //Copy to window view model.
            foreach (var workbookInfo in workbookInfos)
                WorkbookInfos.Add(workbookInfo);

            ShowProgressBar = false;                        
        });

        //On select changed of the first listview, update the datasource of the second listview.
        public ICommand WorkbookInfoSelectionChanged => new DelegateCommand((o) =>
        {
            if (SelectedWorkbookInfo != null)
                SelectedWorksheetInfo = SelectedWorkbookInfo.WorksheetInfos.ElementAt(0);
        });

        //Open export folder
        public ICommand OpenExportFolderCommand => new DelegateCommand(async (o) =>
        {
            if (SelectedWorkbookInfo == null)
                return;
            var folderPath = SelectedWorkbookInfo.FolderPath;
            var fileName = SelectedWorkbookInfo.FileNameNoExtension;
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
                if (WorkbookInfos.Any(x => x.FilePath == filePath))
                    continue;
                else
                {
                    var workbookInfo = new WorkbookInfo { FilePath = filePath };
                    foreach (var sheetName in CommonMethods.GetWorksheetNames(filePath))
                        workbookInfo.WorksheetInfos.Add(new WorksheetInfo { SheetName = sheetName });

                    WorkbookInfos.Add(workbookInfo);
                }
            }
        });

        public ICommand DeleteFilePathCommand => new DelegateCommand((o) =>
        {
            if (SelectedWorkbookInfo == null)
                return;
            WorkbookInfos.Remove(SelectedWorkbookInfo);
        });

        public ICommand SingleSheetCheckedCommand => new DelegateCommand((o) =>
        {
            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand SelectByKeywordCommand => new DelegateCommand((o) =>
        {
            if (Keyword.IsNullOrEmpty())
                return;
            if (SelectedWorkbookInfo == null)
                return;
            foreach (var worksheetInfo in SelectedWorkbookInfo.WorksheetInfos)
                if (worksheetInfo.SheetName.ToUpper().Contains(Keyword.ToUpper()))
                    worksheetInfo.IsSheetChecked = true;

            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand UnSelectByKeywordCommand => new DelegateCommand((o) =>
        {
            if (Keyword.IsNullOrEmpty())
                return;
            if (SelectedWorkbookInfo == null)
                return;
            foreach (var worksheetInfo in SelectedWorkbookInfo.WorksheetInfos)
                if (worksheetInfo.SheetName.ToUpper().Contains(Keyword.ToUpper()))
                    worksheetInfo.IsSheetChecked = false;

            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand SelectAllCommand => new DelegateCommand((o) =>
        {
            foreach (var workbookInfo in WorkbookInfos)
                foreach (var worksheetInfo in workbookInfo.WorksheetInfos)
                    worksheetInfo.IsSheetChecked = true;

            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand UnSelectAllCommand => new DelegateCommand((o) =>
        {
            foreach (var workbookInfo in WorkbookInfos)
                foreach (var worksheetInfo in workbookInfo.WorksheetInfos)
                    worksheetInfo.IsSheetChecked = false;

            NotifyPropertyChanged(nameof(SelectedSheetMessage));
        });

        public ICommand ExportSettingCommand => new DelegateCommand(async (O) =>
        {
            var vm = new ExcelToPaperExportSelectionViewModel();
            //Init old value.
            vm.ExportToSingleFolder = mPrintSettings.ExportToSingleFolder;
            vm.ExportToSeparateFolder = mPrintSettings.ExportToSeparateFolder;

            vm.SingleFolderPath = mPrintSettings.SingleFolderPath;
            vm.AttachWorkbookNameBeforeWorksheet = mPrintSettings.AttachWorkbookNameBeforeWorksheet;

            vm.PrintToPaper = mPrintSettings.PrintToPaper;
            vm.PrintToPdf = mPrintSettings.PrintToPdf;

            vm.MergeNothing = mPrintSettings.MergeNothing;
            vm.MergeToFileSeparately = mPrintSettings.MergeToFileSeparately;
            vm.MergeToSingleFile = mPrintSettings.MergeToSingleFile;
            vm.MergeDeleteOriginFile = mPrintSettings.MergeDeleteOriginFile;
            //Show dialog.
            var result = await DialogHost.Show(vm.View, "Root");
            var dialogResult = (bool)result;
            if (!dialogResult) return;

            mPrintSettings.ExportToSingleFolder = vm.ExportToSingleFolder;
            mPrintSettings.ExportToSeparateFolder = vm.ExportToSeparateFolder;

            mPrintSettings.AttachWorkbookNameBeforeWorksheet = vm.AttachWorkbookNameBeforeWorksheet;
            mPrintSettings.SingleFolderPath = vm.SingleFolderPath;

            mPrintSettings.PrintToPaper = vm.PrintToPaper;
            mPrintSettings.PrintToPdf = vm.PrintToPdf;

            mPrintSettings.MergeNothing = vm.MergeNothing;
            mPrintSettings.MergeToFileSeparately = vm.MergeToFileSeparately;
            mPrintSettings.MergeToSingleFile = vm.MergeToSingleFile;
            mPrintSettings.MergeDeleteOriginFile = vm.MergeDeleteOriginFile;
        });

        
        public ICommand GetPageCountSizeCommand => new DelegateCommand(async (o) =>
        {
            if (ExcelIsBusy)
            {
                UpdateProgressMessage("Excelは忙しいです。");
                await Task.Delay(2000);
                UpdateProgressMessage("");
                return;
            }

            //Show dialog
            var vm = new ExcelToPaperGetDetailSelectionViewModel();
            var result = await DialogHost.Show(vm.View, "MiddleRight");
            var dialogResult = (bool)result;
            if (!dialogResult) return;

            //Show progress bar
            ShowPreviewProgressBar = true;

            ExcelIsBusy = true;
            //Start a excel
            await StartExcel();

            if (Excel != null)
            {
                if (!vm.GetAllFileDetail)
                {
                    if (SelectedWorkbookInfo == null || SelectedWorkbookInfo.FilePath.IsNullOrEmpty())
                        return;
                    await CommonMethods.GetWorksheetPageCount(Excel, SelectedWorkbookInfo.FilePath, SelectedWorkbookInfo.WorksheetInfos, UpdateProgressMessage);
                }
                else
                    foreach (var workbookInfo in WorkbookInfos)
                        await CommonMethods.GetWorksheetPageCount(Excel, workbookInfo.FilePath, workbookInfo.WorksheetInfos, UpdateProgressMessage);
            }
            ExcelIsBusy = false;
            ShowPreviewProgressBar = false;
        });
        

        public ICommand PreviewCommand => new DelegateCommand(async (o) =>
        {
            if (ExcelIsBusy)
            {
                UpdateProgressMessage("Excelは忙しいです。");
                await Task.Delay(2000);
                UpdateProgressMessage("");
                return;
            }

            //Show progress bar
            ShowPreviewProgressBar = true;

            if (Excel == null)
                return;

            if (SelectedWorkbookInfo == null || SelectedWorkbookInfo.FilePath.IsNullOrEmpty())
                return;


            ExcelIsBusy = true;
            //Start a excel
            await StartExcel();

            await CommonMethods.GetWorkSheetPreview(Excel, SelectedWorkbookInfo.FilePath, SelectedWorkbookInfo.WorksheetInfos, UpdateProgressMessage);
            foreach (var si in SelectedWorkbookInfo)
                si.UpdatePreviews();
            ExcelIsBusy = false;

            ShowPreviewProgressBar = false;
        });

        public ICommand ExcelToPaperCancelCommand => new DelegateCommand((o) =>
        {
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

            //Wait until export finish.
            var printResults = await ExcelPrintMethods.PrintToPaper(
                SelectedPrinter,
                mCancelTokenSource.Token,
                WorkbookInfos,
                mPrintSettings,
                UpdateProgressMessage);

            //Wait 1.5s to let the last printed pdf to be released.
            await Task.Delay(1500);

            //Merge pdf
            if (mPrintSettings.MergeToFileSeparately)
            {
                foreach (var printResult in printResults)
                {
                    PdfMethods.MergePdf(printResult, UpdateProgressMessage);
                }
            }

            //Delete merge pdf
            if (mPrintSettings.MergeDeleteOriginFile)
            {
                foreach (var printResult in printResults)
                {
                    PdfMethods.DeletePdf(printResult.PrintedPdfPaths);
                }
            }

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
        });

        public ICommand CancelCommand => new DelegateCommand((o) =>
        {
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

        private async Task GetPageCountAndSize(Application excel, IEnumerable<WorkbookInfo> workbookInfos, Action<string> updateStatus=null)
        {
            foreach (var workbookInfo in workbookInfos)
            {
                if (workbookInfo.IsWorksheetPageCountSizeObtained)
                    continue;
                await CommonMethods.GetWorksheetPageCount(excel, workbookInfo.FilePath, workbookInfo.WorksheetInfos, updateStatus);
            }
        }

        private async void GetWorksheetPreview(Application excel, WorkbookInfo workbookInfo, Action<string> updateStatus = null)
        {
            if (workbookInfo.IsWorksheetPreviewObtained)
                return;
            await CommonMethods.GetWorkSheetPreview(excel, workbookInfo.FilePath, workbookInfo.WorksheetInfos, updateStatus);
        }

        private async Task StartExcel()
        {
            if (Excel == null)
                await Task.Run(() =>
                {
                    Excel = new Application();
                    Excel.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                });
        }
    }
}