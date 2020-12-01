using ExcelToPaper.DataModels;
using Microsoft.Office.Interop.Excel;
using System;
using System.Threading;
using System.Threading.Tasks;


namespace ExcelToPaper.Components
{
    //This class gets preview of worksheets inside a workbook
    internal class WorkbookPreviewExtractor
    {
        private static Application ExcelPreview { get; set; }
        public CancellationToken CancellationToken { get; private set; }
        //Update status function injected from outside
        public Action<string> UpdateStatus { get; set; }

        public WorkbookPreviewExtractor(CancellationToken cancellationToken)
        {
            CancellationToken = cancellationToken;
        }

        ~WorkbookPreviewExtractor()
        {
            if (ExcelPreview != null)
            {
                ExcelPreview.Workbooks.Close();
                ExcelPreview.Quit();
                ExcelPreview = null;
            }
        }

        public void SetCancelToken(CancellationToken cancellationToken)
        {
            CancellationToken = cancellationToken;
        }

        //Glue method
        public async Task GetPagePreview(WorkbookInfo workbookInfo)
        {
            await StartExcelPreview();
            await GetWorksheetPreview(ExcelPreview, workbookInfo);
            UpdateWorksheetPreview(workbookInfo);
        }

        private async Task StartExcelPreview()
        {
            if (ExcelPreview == null)
            {
                await Task.Run(() =>
                {
                    ExcelPreview = new Application();
                    ExcelPreview.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                });
            }
        }

        private async Task GetWorksheetPreview(Application excel, WorkbookInfo workbookInfo)
        {
            if (workbookInfo.IsWorksheetPreviewObtained)
                return;
            await CommonMethods.GetWorkSheetPreview(excel, workbookInfo.FilePath, workbookInfo.WorksheetInfos, CancellationToken, UpdateStatus);
            workbookInfo.IsWorksheetPreviewObtained = true;
        }

        private void UpdateWorksheetPreview(WorkbookInfo workbookInfo)
        {
            foreach (var si in workbookInfo)
                si.UpdatePreviews();
        }
    }
}
