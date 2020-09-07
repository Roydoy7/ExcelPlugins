using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;

namespace ExcelToPdf.DataConverters
{
    class PathToFileNameDataConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var filePath = value as string;
            if (filePath != null)
                return Path.GetFileName(filePath);
            return "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
