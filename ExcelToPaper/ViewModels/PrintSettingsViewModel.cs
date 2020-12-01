using CommonTools.FileBrowsers;
using CommonWPFTools;
using ExcelToPaper.DataModels;
using ExcelToPaper.Views;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelToPaper.ViewModels
{
    internal class PrintSettingsViewModel : ViewModelBase<PrintSettingsView>
    {
        public PrintSettings PrintSettings { get; private set; } = new PrintSettings();
        public PrintSettingsViewModel(PrintSettings printSettings)
        {
            PrintSettings.Copy(printSettings);
        }

        public bool ShowErrorMessage { get; set; }
        public string ErrorMessage { get; set; }


        public ICommand OpenFolderCommand => new DelegateCommand((o) =>
        {
            PrintSettings.SingleFolderPath = FileBrowser.BrowseFolder();
        });

        public ICommand CloseCommand => new DelegateCommand(async (o) =>
        {

            //Check if the path is valid.
            if (PrintSettings.ExportToSingleFolder && !PrintSettings.PrintToPaper)
                if (!Directory.Exists(PrintSettings.SingleFolderPath))
                {
                    ErrorMessage = "パスが存在しません。";
                    ShowErrorMessage = true;
                    await Task.Delay(2000);
                    ErrorMessage = "";
                    ShowErrorMessage = false;
                    return;
                }

            View.Close();
            //DialogHost.CloseDialogCommand.Execute(true, null);
        });

    }
}