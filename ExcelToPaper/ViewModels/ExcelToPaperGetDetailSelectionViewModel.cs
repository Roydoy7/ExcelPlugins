using CommonWPFTools;
using ExcelToPaper.Views;

namespace ExcelToPaper.ViewModels
{
    internal class ExcelToPaperGetDetailSelectionViewModel : ViewModelBase<ExcelToPaperGetDetailSelectionView>
    {
        public bool GetAllFileDetail { get; set; } = false;
    }
}