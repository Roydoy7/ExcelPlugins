using CommonTools.FileBrowsers;
using CommonWPFTools;
using ExcelToPdf.Views;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace ExcelToPdf.ViewModels
{
    class ExcelToPdfFormViewModel : ViewModelBase<ExcelToPdfForm>
    {
        public string ExcelFolderPath { get; set; }
        public ICommand OpenExcelCommand => new DelegateCommand((o) =>
        {
            ExcelFolderPath = FileBrowser.BrowseFolder();
        });
        public ICommand OkCommand => new DelegateCommand((o) =>
        {
            if (!Directory.Exists(ExcelFolderPath))
            {
                MessageBox.Show("パスが正しくありません。", "Error", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            View.DialogResult = true;
            View.Close();
        });
    }
}
