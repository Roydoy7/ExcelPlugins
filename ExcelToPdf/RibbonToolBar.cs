using ExcelToPdf.Commands;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelToPdf
{
    public partial class RibbonToolBar
    {
        private void RibbonToolBar_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void BatchToPdf_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCommands.BatchToPdf();
        }
    }
}
