using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using PropertyChanged;

namespace ExcelToPaper.DataModels
{
    public class WorkbookInfo : INotifyPropertyChanged, IEnumerable<WorksheetInfo>
    {
        public Action WorksheetChecked { get; set; }
        public bool IsWorksheetPageCountSizeObtained { get; set; } = false;
        public bool IsWorksheetPreviewObtained { get; set; } = false;

        [OnChangedMethod(nameof(OnWorksheetChecked))]
        public bool? IsAllWorksheetChecked
        {
            get
            {
                if (WorksheetInfos.Any(x => x.IsWorksheetChecked) && WorksheetInfos.Any(x => !x.IsWorksheetChecked))
                {
                    IsThreeState = true;
                    return null;
                }
                IsThreeState = false;
                if (WorksheetInfos.Any(x => x.IsWorksheetChecked))
                    return true;
                else
                    return false;
            }
            set
            {
                if (value.Value)
                {
                    foreach (var si in WorksheetInfos)
                        si.IsWorksheetChecked = true;
                }
                else
                    foreach (var si in WorksheetInfos)
                        si.IsWorksheetChecked = false;

                WorksheetChecked?.Invoke();
            }
        }
        public bool IsThreeState { get; set; }
        public bool ShowProgressBar { get; set; } = false;
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
        public ObservableCollection<WorksheetInfo> WorksheetInfos { get; private set; } 
            = new ObservableCollection<WorksheetInfo>();

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

        public void Add(WorksheetInfo worksheetInfo)
        {
            worksheetInfo.WorksheetChecked = OnWorksheetChecked;
            WorksheetInfos.Add(worksheetInfo);
        }

        public void OnWorksheetChecked()
        {
            WorksheetChecked?.Invoke();
            NotifyPropertyChanged(nameof(IsAllWorksheetChecked));
        }
    }
}
