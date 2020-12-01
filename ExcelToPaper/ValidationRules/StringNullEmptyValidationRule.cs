using System.Globalization;
using System.Windows.Controls;

namespace ExcelToPaper.ValidationRules
{
    internal class StringNullEmptyValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            var str = (string)value;
            if (str == null || str.Length == 0)
            {
                return new ValidationResult(false, "Content is null or empty.");
            }
            return new ValidationResult(true, null);
        }
    }
}
