﻿using ExcelToPaper.Components;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;

namespace ExcelToPaper.DataModels
{
    public class WorksheetInfo : INotifyPropertyChanged
    {
        public string SheetName { get; set; }
        public bool IsSheetChecked { get; set; } = false;
        public int Count { get; set; } = 0;
        public uint StartPage { get; set; } = 0;
        public uint EndPage { get; set; } = 0;
        public XlPaperSize PaperSize { get; set; }
        public XlPageOrientation Orientation { get; set; }
        public List<Bitmap> PreviewsRaw { get; private set; } = new List<Bitmap>();
        public ObservableCollection<PreviewInfo> Previews { get; private set; } = new ObservableCollection<PreviewInfo>();

        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public void UpdatePreviews()
        {
            Previews.Clear();
            var index = 1;
            foreach (var bmp in PreviewsRaw)
                Previews.Add(new PreviewInfo
                {
                    Preview = bmp.ToImageSource(),
                    Index = index++,
                });
        }
    }
}
