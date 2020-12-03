using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelToPaper.DataModels
{
    internal class WorkbookCollection : ObservableCollection<WorkbookInfo>, INotifyPropertyChanged
    {
        public WorkbookCollection()
        {
            base.CollectionChanged += (o, e) =>
            {
                OnWorksheetChecked();
            };
        }
        public int SelectedWorksheetCount
        {
            get
            {
                var count = 0;
                foreach (var workbookInfo in this)
                    foreach (var worksheetInfo in workbookInfo)
                        if (worksheetInfo.IsWorksheetChecked)
                            count++;
                return count;
            }
        }
        public int SelectedPageCount
        {
            get
            {
                var count = 0;
                foreach (var workbookInfo in this)
                    foreach (var worksheetInfo in workbookInfo)
                        if (worksheetInfo.IsWorksheetChecked)
                            count += worksheetInfo.Count;
                return count;
            }
        }

        public new event PropertyChangedEventHandler PropertyChanged;

        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public new void Add(WorkbookInfo workbookInfo)
        {
            workbookInfo.WorksheetChecked = OnWorksheetChecked;
            base.Add(workbookInfo);
        }

        public void OnWorksheetChecked()
        {
            NotifyPropertyChanged(nameof(SelectedWorksheetCount));
            NotifyPropertyChanged(nameof(SelectedPageCount));
        }
    }
}
