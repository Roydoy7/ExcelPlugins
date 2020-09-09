﻿using CommonWPFTools;
using System.IO;
using ExcelToPaper.Views;
using System.Windows.Input;
using MaterialDesignThemes.Wpf;
using System.Threading.Tasks;
using CommonTools.FileBrowsers;
using ExcelToPaper.DataModels;

namespace ExcelToPaper.ViewModels
{
    internal class ExcelToPaperExportSelectionViewModel : ViewModelBase<ExcelToPaperExportSelectionView>
    {
        public PrintSettings PrintSettings { get; set; } = new PrintSettings();
        //Is print to paper.
        public bool PrintToPaper
        {
            get { return PrintSettings.PrintToPaper; }
            set { PrintSettings.PrintToPaper = value; }
        }
        //Export to a single folder
        public bool ExportToSingleFolder
        {
            get { return PrintSettings.ExportToSingleFolder; }
            set { PrintSettings.ExportToSingleFolder = value; }
        }
        //Attach workbook name before work sheet name
        public bool AttachWorkbookNameBeforeWorksheet
        {
            get { return PrintSettings.AttachWorkbookNameBeforeWorksheet; }
            set { PrintSettings.AttachWorkbookNameBeforeWorksheet = value; }
        }
        //The folder path that pdf will be printed to.
        public string SingleFolderPath
        {
            get { return PrintSettings.SingleFolderPath; }
            set { PrintSettings.SingleFolderPath = value; }
        }
        public string ErrorMessage { get; set; }
        public ICommand OpenFolderCommand => new DelegateCommand((o) => {
            SingleFolderPath = FileBrowser.BrowseFolder();
            NotifyPropertyChanged(nameof(SingleFolderPath));
        });

        public ICommand OkCommand => new DelegateCommand(async(o) => {

            //Check if the path is valid.
            if(ExportToSingleFolder&&!PrintToPaper)
                if(!Directory.Exists(SingleFolderPath))
                {
                    ErrorMessage = "パスが存在しません。";
                    NotifyPropertyChanged(nameof(ErrorMessage));
                    await Task.Delay(2000);
                    ErrorMessage = "";
                    NotifyPropertyChanged(nameof(ErrorMessage));
                    return;
                }

            DialogHost.CloseDialogCommand.Execute(true, null);
        });
    }
}