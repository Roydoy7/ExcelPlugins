using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelToPaper.DataModels
{
    public class WorkbookInfo: INotifyPropertyChanged, IEnumerable<WorksheetInfo>
    {
        public bool IsWorksheetPageCountSizeObtained { get; set; } = false;
        public bool IsWorksheetPreviewObtained { get; set; } = false;
        public string FilePath { get; set; }        
        public string FileName
        {
            get => Path.GetFileName(FilePath);
        }
        public string FileNameNoExtension
        {
            get => Path.GetFileNameWithoutExtension(FilePath);
        }
        public string FolderPath
        {
            get => Path.GetDirectoryName(FilePath);
        }
        public ObservableCollection<WorksheetInfo> WorksheetInfos { get; private set; } = new ObservableCollection<WorksheetInfo>();
        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public IEnumerator<WorksheetInfo> GetEnumerator()
        {
            return WorksheetInfos.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
