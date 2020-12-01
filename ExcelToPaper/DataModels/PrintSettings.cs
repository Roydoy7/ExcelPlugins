using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelToPaper.DataModels
{
    public class PrintSettings:INotifyPropertyChanged
    {
        //Is print to paper.
        public bool PrintToPaper { get; set; }
        public bool PrintToPdf { get; set; } = true;

        //Export to a separate folders
        public bool ExportToSeparateFolder { get; set; } = true;
        public bool ExportToSingleFolder { get; set; } = false;

        //Attach workbook name before work sheet name
        public bool AttachWorkbookNameBeforeWorksheet { get; set; } = true;

        //The single folder path that pdf will be printed to.
        public string SingleFolderPath { get; set; }
        //Don't merge
        public bool MergeNothing { get; set; } 
        //Merge to file separately
        public bool MergeToFileSeparately { get; set; } = true;
        //Merge to all to a single file
        public bool MergeToSingleFile { get; set; }
        public bool MergeDeleteOriginFile { get; set; } = false;

        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        public void NofityAllPropertyChanged()
        {
            foreach (var p in this.GetType().GetProperties())
                NotifyPropertyChanged(p.Name);
        }
        public void Copy(PrintSettings printSettings)
        {
            PrintToPaper = printSettings.PrintToPaper;
            PrintToPdf = printSettings.PrintToPdf;

            ExportToSingleFolder = printSettings.ExportToSingleFolder;
            ExportToSeparateFolder = printSettings.ExportToSeparateFolder;

            AttachWorkbookNameBeforeWorksheet = printSettings.AttachWorkbookNameBeforeWorksheet;

            SingleFolderPath = printSettings.SingleFolderPath;

            MergeNothing = printSettings.MergeNothing;
            MergeToFileSeparately = printSettings.MergeToFileSeparately;
            MergeToSingleFile = printSettings.MergeToSingleFile;
            MergeDeleteOriginFile = printSettings.MergeDeleteOriginFile;
        }
    }
}