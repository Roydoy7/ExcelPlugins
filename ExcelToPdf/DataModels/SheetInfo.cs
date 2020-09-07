using System.ComponentModel;

namespace ExcelToPdf.DataModels
{
    internal class SheetInfo:INotifyPropertyChanged
    {
        public string SheetName { get; set; }
        public bool IsSheetChecked { get; set; } = false;
        public int Count { get; set; } = 0;

        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
