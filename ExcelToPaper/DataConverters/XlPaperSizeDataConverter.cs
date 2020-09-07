using Microsoft.Office.Interop.Excel;
using System;
using System.Globalization;
using System.Windows.Data;

namespace ExcelToPaper.DataConverters
{
    class XlPaperSizeDataConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var paperSize = (XlPaperSize)value;
            if (paperSize == 0)
                return "";
            return paperSize.ToString("g").Substring(7);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
