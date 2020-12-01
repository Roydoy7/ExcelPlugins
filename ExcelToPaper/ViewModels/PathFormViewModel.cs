using CommonTools.FileBrowsers;
using CommonWPFTools;
using ExcelToPaper.Views;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelToPaper.ViewModels
{
    internal class PathFormViewModel : ViewModelBase<PathForm>
    {
        public string ExcelFolderPath { get; set; }
        public bool ShowErrorMessage { get; set; }
        public string ErrorMessage { get; set; }

        public ICommand OpenExcelCommand => new DelegateCommand((o) =>
        {
            ExcelFolderPath = FileBrowser.BrowseFolder();
            NotifyPropertyChanged(nameof(ExcelFolderPath));
        });
        public ICommand OkCommand => new DelegateCommand(async (o) =>
        {
            if (!Directory.Exists(ExcelFolderPath))
            {
                ErrorMessage = "パスが存在しません。";
                ShowErrorMessage = true;
                await Task.Delay(2000);
                ErrorMessage = "";
                ShowErrorMessage = false;
                return;
            }
            View.DialogResult = true;
            View.Close();
        });
    }
}
