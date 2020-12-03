using CommonTools;
using CommonTools.FileBrowsers;
using CommonWPFTools;
using ExcelToPaper.Components;
using ExcelToPaper.DataModels;
using ExcelToPaper.Views;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Threading;
using Win = System.Windows;

namespace ExcelToPaper.ViewModels
{
    internal class ExcelToPaperFormViewModel : ViewModelBase<ExcelToPaperForm>
    {
        #region Fileds
        private bool mSelectAllSheet;
        #endregion

        #region Constructors
        public ExcelToPaperFormViewModel()
        {
            //Set default printer
            var printSettings = new PrinterSettings();
            SelectedPrinter = printSettings.PrinterName;

            //
            PageSizeCountExtractor = new WorkbookPageSizeCountExtractor(CancelTokenSourceOther.Token);
            PageSizeCountExtractor.UpdateStatus = UpdateProgressMessage;
            //
            PagePreviewExtractor = new WorkbookPreviewExtractor(CancelTokenSourceOther.Token);
            PagePreviewExtractor.UpdateStatus = UpdateProgressMessage;

        }
        #endregion

        #region Properties
        public static Application ExcelPrint { get; private set; }

        public bool SelectAllSheet
        {
            get => mSelectAllSheet;
            set
            {
                mSelectAllSheet = value;

                if (SelectedWorkbookInfo != null)
                    foreach (var worksheetInfo in SelectedWorkbookInfo.WorksheetInfos)
                        worksheetInfo.IsWorksheetChecked = value;
            }
        }
        public bool ShowPreviewProgressBar { get; set; }
        public bool ShowProgressBar { get; set; }
        public int SheetCount { get; set; }
        public int PageCount { get; set; }
        public string Keyword { get; set; } = "";
        public string SelectedPrinter { get; set; }
        public string PrintButtonBadge { get; set; } = "Pdf";
        public string ProgressMessage { get; set; }

        public CancellationTokenSource CancelTokenSourcePrint { get; private set; }
        public CancellationTokenSource CancelTokenSourceOther { get; private set; } = new CancellationTokenSource();
        public PrintSettings PrintSettings { get; private set; } = new PrintSettings();
        public Task GetPageCountSizeTask { get; private set; }
        public WorkbookPageSizeCountExtractor PageSizeCountExtractor { get; private set; }
        public WorkbookPreviewExtractor PagePreviewExtractor { get; private set; }
        public WorkbookInfo SelectedWorkbookInfo { get; set; }
        public WorksheetInfo SelectedWorksheetInfo { get; set; }
        public WorkbookCollection WorkbookInfos { get; set; } = new WorkbookCollection();

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
        #endregion

        #region Commands

        public ICommand OnWindowClosing => new DelegateCommand((o) =>
        {
            CancelTokenSourceOther.Cancel();
        });


        public ICommand AddFromFolderCommand => new DelegateCommand(async (o) =>
        {
            var vm = new PathFormViewModel();
            vm.View.ShowDialog();
            if (vm.ExcelFolderPath.IsNullOrEmpty())
                return;

            //Create a new cancel token if old token has cancelled
            if (CancelTokenSourceOther == null || CancelTokenSourceOther.IsCancellationRequested)
            {
                CancelTokenSourceOther = new CancellationTokenSource();
                PageSizeCountExtractor.SetCancelToken(CancelTokenSourceOther.Token);
                PagePreviewExtractor.SetCancelToken(CancelTokenSourceOther.Token);
            }

            var folderPath = vm.ExcelFolderPath;

            //Show progress bar
            ShowProgressBar = true;

            //Get file paths
            var filePaths = CoreMethods.GetExcelPath(folderPath);
            //Create workbookinfos
            var newWorkbookInfos = CreateWorkbookInfo(filePaths);
            //Exclude existing filepaths
            var extractedWorkbookInfos = ExtractNewWorkbookInfos(WorkbookInfos, newWorkbookInfos);

            //Add to UI first
            foreach (var workbookInfo in extractedWorkbookInfos)
                WorkbookInfos.Add(workbookInfo);

            //Get worksheet names
            await GetWorksheetName(extractedWorkbookInfos);

            //Hide progress bar
            ShowProgressBar = false;

            //Read page count and size
            PageSizeCountExtractor.GetPageCountSize(extractedWorkbookInfos);
        });

        public ICommand AddFilePathCommand => new DelegateCommand(async (o) =>
        {
            var filePaths = new List<string>();
            foreach (var filePath in FileBrowser.BrowseExcelFile(true))
                filePaths.Add(filePath);
            //If no files selected, return
            if (!filePaths.Any())
                return;

            //Create a new cancel token if old token has cancelled
            if (CancelTokenSourceOther == null || CancelTokenSourceOther.IsCancellationRequested)
            {
                CancelTokenSourceOther = new CancellationTokenSource();
                PageSizeCountExtractor.SetCancelToken(CancelTokenSourceOther.Token);
                PagePreviewExtractor.SetCancelToken(CancelTokenSourceOther.Token);
            }

            //Show progress bar
            ShowProgressBar = true;

            //Create workbookinfos
            var newWorkbookInfos = CreateWorkbookInfo(filePaths);
            //Exclude existing filepaths
            var extractedWorkbookInfos = ExtractNewWorkbookInfos(WorkbookInfos, newWorkbookInfos);

            //Add to UI first
            foreach (var workbookInfo in extractedWorkbookInfos)
                WorkbookInfos.Add(workbookInfo);

            //Get worksheet names
            await GetWorksheetName(extractedWorkbookInfos);

            //Show progress bar
            ShowProgressBar = false;

            //Read page count and size
            PageSizeCountExtractor.GetPageCountSize(extractedWorkbookInfos);
        });

        //On select changed of the first listview, update the datasource of the second listview.
        public ICommand WorkbookInfoSelectionChanged => new DelegateCommand((o) =>
        {
            if (SelectedWorkbookInfo != null && SelectedWorkbookInfo.WorksheetInfos.Count != 0)
            {
                SelectedWorksheetInfo = SelectedWorkbookInfo.WorksheetInfos.ElementAt(0);
                if (SelectedWorkbookInfo.Any(x => x.IsWorksheetChecked))
                    mSelectAllSheet = true;
                else
                    mSelectAllSheet = false;
                NotifyPropertyChanged(nameof(SelectAllSheet));
            }
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

        public ICommand DeleteFilePathCommand => new DelegateCommand((o) =>
        {
            if (SelectedWorkbookInfo == null)
                return;
            WorkbookInfos.Remove(SelectedWorkbookInfo);
        });

        public ICommand ClearFilePathCommand => new DelegateCommand((o) =>
        {
            CancelTokenSourceOther.Cancel();
            WorkbookInfos.Clear();
        });

        public ICommand SelectByKeywordCommand => new DelegateCommand((o) =>
        {
            if (Keyword.IsNullOrEmpty())
                return;
            if (SelectedWorkbookInfo == null)
                return;
            foreach (var worksheetInfo in SelectedWorkbookInfo.WorksheetInfos)
                if (worksheetInfo.SheetName.ToUpper().Contains(Keyword.ToUpper()))
                    worksheetInfo.IsWorksheetChecked = true;
        });

        public ICommand SelectAllByKeywordCommand => new DelegateCommand((o) =>
        {
            if (Keyword.IsNullOrEmpty())
                return;

            foreach (var workbookInfo in WorkbookInfos)
                foreach (var worksheetInfo in workbookInfo)
                    if (worksheetInfo.SheetName.ToUpper().Contains(Keyword.ToUpper()))
                        worksheetInfo.IsWorksheetChecked = true;
        });

        public ICommand UnSelectByKeywordCommand => new DelegateCommand((o) =>
        {
            if (Keyword.IsNullOrEmpty())
                return;
            if (SelectedWorkbookInfo == null)
                return;
            foreach (var worksheetInfo in SelectedWorkbookInfo.WorksheetInfos)
                if (worksheetInfo.SheetName.ToUpper().Contains(Keyword.ToUpper()))
                    worksheetInfo.IsWorksheetChecked = false;
        });

        public ICommand UnSelectAllByKeywordCommand => new DelegateCommand((o) =>
        {
            if (Keyword.IsNullOrEmpty())
                return;

            foreach (var workbookInfo in WorkbookInfos)
                foreach (var worksheetInfo in workbookInfo)
                    if (worksheetInfo.SheetName.ToUpper().Contains(Keyword.ToUpper()))
                        worksheetInfo.IsWorksheetChecked = false;
        });

        public ICommand SelectAllCommand => new DelegateCommand((o) =>
        {
            foreach (var workbookInfo in WorkbookInfos)
                foreach (var worksheetInfo in workbookInfo.WorksheetInfos)
                    worksheetInfo.IsWorksheetChecked = true;
        });

        public ICommand UnSelectAllCommand => new DelegateCommand((o) =>
        {
            foreach (var workbookInfo in WorkbookInfos)
                foreach (var worksheetInfo in workbookInfo.WorksheetInfos)
                    worksheetInfo.IsWorksheetChecked = false;
        });

        public ICommand MoveWorksheetUp => new DelegateCommand((o) =>
        {
            if (SelectedWorkbookInfo == null)
                return;

            if (SelectedWorksheetInfo == null)
                return;

            var cnt = SelectedWorkbookInfo.WorksheetInfos.Count;
            var index = SelectedWorkbookInfo.WorksheetInfos.IndexOf(SelectedWorksheetInfo);
            if (index == 0) return;
            SelectedWorkbookInfo.WorksheetInfos.Move(index, index - 1);
        });

        public ICommand MoveWorksheetDown => new DelegateCommand((o) =>
        {
            if (SelectedWorkbookInfo == null)
                return;

            if (SelectedWorksheetInfo == null)
                return;

            var cnt = SelectedWorkbookInfo.WorksheetInfos.Count;
            var index = SelectedWorkbookInfo.WorksheetInfos.IndexOf(SelectedWorksheetInfo);
            if (index == cnt - 1) return;
            SelectedWorkbookInfo.WorksheetInfos.Move(index, index + 1);
        });

        public ICommand ExportSettingCommand => new DelegateCommand((O) =>
        {
            var vm = new PrintSettingsViewModel(PrintSettings);
            vm.View.ShowDialog();
            PrintSettings.Copy(vm.PrintSettings);
            PrintSettings.NofityAllPropertyChanged();
            UpdatePrintButtonBadge();
        });

        public ICommand PreviewCommand => new DelegateCommand(async (o) =>
        {
            var curWorkbookInfo = SelectedWorkbookInfo;
            if (curWorkbookInfo == null || curWorkbookInfo.FilePath.IsNullOrEmpty())
                return;

            //If in the progress of getting preview
            if (ShowPreviewProgressBar)
                return;

            //If page count is very large, ask if continue.
            if (curWorkbookInfo.WorksheetInfos.Select(x => x.Count).Sum() > 35)
            {
                if (Win.MessageBox.Show("ページの数が多いので、時間がかかります。続きますか？", "Question", Win.MessageBoxButton.YesNo, Win.MessageBoxImage.Question) != Win.MessageBoxResult.Yes)
                    return;
            }

            //Show progress bar
            ShowPreviewProgressBar = true;
            curWorkbookInfo.ShowProgressBar = true;
            //Start a excel
            await PagePreviewExtractor.GetPagePreview(curWorkbookInfo);
            curWorkbookInfo.ShowProgressBar = false;
            ShowPreviewProgressBar = false;
        });

        public ICommand ExcelToPaperCancelCommand => new DelegateCommand((o) =>
        {
            if (CancelTokenSourcePrint == null)
                return;
            CancelTokenSourcePrint.Cancel();
        });

        public ICommand OkCommand => new DelegateCommand(async (o) =>
        {
            //Check the cancel token source.
            if (CancelTokenSourcePrint != null)
                return;
            //Show progress bar
            ShowProgressBar = true;

            //Init cancel token.
            CancelTokenSourcePrint = new CancellationTokenSource();

            //Wait until export finish.
            var printResults = await ExcelPrintMethods.PrintToPaper(
                SelectedPrinter,
                CancelTokenSourcePrint.Token,
                WorkbookInfos,
                PrintSettings,
                UpdateProgressMessage);

            //Wait 1.5s to let the last printed pdf to be released.
            await Task.Delay(1500);

            //Merge pdf
            if (PrintSettings.MergeToFileSeparately)
            {
                foreach (var printResult in printResults)
                    PdfMethods.MergePdf(printResult, UpdateProgressMessage);
            }

            //Delete merge pdf
            if (PrintSettings.MergeToFileSeparately && PrintSettings.MergeDeleteOriginFile)
            {
                foreach (var printResult in printResults)
                    PdfMethods.DeletePdf(printResult.PrintedPdfPaths);
            }

            //Set finish message
            ProgressMessage = "完成";

            //Dispose cancel token source
            CancelTokenSourcePrint.Dispose();
            CancelTokenSourcePrint = null;

            //Wait 5s and clear finish message, hide progress bar
            await Task.Delay(5000);
            ProgressMessage = "";

            //Hide progressbar
            ShowProgressBar = false;
        });
        #endregion

        #region Private methods
        private IEnumerable<WorkbookInfo> CreateWorkbookInfo(IEnumerable<string> filePaths)
        {
            var workbookInfos = new List<WorkbookInfo>();
            foreach (var filePath in filePaths)
                workbookInfos.Add(new WorkbookInfo { FilePath = filePath });
            return workbookInfos;
        }
        private IEnumerable<WorkbookInfo> ExtractNewWorkbookInfos(IList<WorkbookInfo> existWorkbookInfos, IEnumerable<WorkbookInfo> newWorkbookInfos)
        {
            var workbookInfos = new List<WorkbookInfo>();
            var existWorkbookInfoDict = existWorkbookInfos.ToDictionary(k => k.FilePath, v => v);
            foreach (var workbookInfo in newWorkbookInfos)
                if (!existWorkbookInfoDict.ContainsKey(workbookInfo.FilePath))
                    workbookInfos.Add(workbookInfo);
            return workbookInfos;
        }
        private async Task<IEnumerable<WorkbookInfo>> GetWorksheetName(IEnumerable<WorkbookInfo> workbookInfos)
        {
            var dispatcher = Dispatcher.CurrentDispatcher;
            await Task.Run(() =>
            {
                Parallel.ForEach(workbookInfos, workbookInfo =>
                {
                    foreach (var sheetName in workbookInfo.GetWorksheetNames())
                        dispatcher.Invoke(() =>
                        {
                            workbookInfo.Add(new WorksheetInfo { WorkbookInfo = workbookInfo, SheetName = sheetName });
                        });
                });
            });

            return workbookInfos;
        }

        private void UpdateProgressMessage(string message)
        {
            ProgressMessage = message;
        }

        private void UpdatePrintButtonBadge()
        {
            if (PrintSettings.PrintToPdf)
                PrintButtonBadge = "Pdf";
            else
                PrintButtonBadge = "Paper";
        }
        #endregion
    }
}