using ExcelToPaper.DataModels;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToPaper.Components
{
    //This class gets preview of worksheets inside a workbook
    internal class WorkbookPreviewExtractor
    {
        #region Properties
        private static Application ExcelPreview { get; set; }
        public CancellationToken CancellationToken { get; private set; }

        //Update status function injected from outside
        public Action<string> UpdateStatus { get; set; }
        #endregion

        #region Constructors
        public WorkbookPreviewExtractor(CancellationToken cancellationToken)
        {
            CancellationToken = cancellationToken;
        }

        ~WorkbookPreviewExtractor()
        {
            if (ExcelPreview != null)
            {
                //ExcelPreview.Workbooks.Close();
                try
                {
                    ExcelPreview.Quit();
                    Marshal.ReleaseComObject(ExcelPreview);
                }
                catch { }
                ExcelPreview = null;
            }
        }
        #endregion

        #region Public methods
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
        #endregion

        #region Private methods

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
            var result = await workbookInfo.GetWorkSheetPreview(excel, CancellationToken, UpdateStatus);
            if (result)
                workbookInfo.IsWorksheetPreviewObtained = true;
        }

        private void UpdateWorksheetPreview(WorkbookInfo workbookInfo)
        {
            foreach (var si in workbookInfo)
                si.UpdatePreviews();
        }
        #endregion
    }
}