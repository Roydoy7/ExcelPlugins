using Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace ExcelToPaper.DataModels
{
    internal class SheetInfo : INotifyPropertyChanged
    {
        public string SheetName { get; set; }
        public bool IsSheetChecked { get; set; } = false;
        public int Count { get; set; } = 0;
        public uint StartPage { get; set; } = 0;
        public uint EndPage { get; set; } = 0;
        public XlPaperSize PaperSize { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
