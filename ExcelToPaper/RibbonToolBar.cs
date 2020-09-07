using ExcelToPaper.Commands;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelToPaper
{
    public partial class RibbonToolBar
    {
        private void RibbonToolBar_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void BatchToPaper_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCommands.BatchPrint();
        }
    }
}
