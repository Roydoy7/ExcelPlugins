using ExcelToPaper.DataModels;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToPaper.Components
{
    //This class is used to obtain worksheets' page count and size inside a workbook
    internal class WorkbookPageSizeCountExtractor
    {
        #region Properties
        //Excel application
        private static Application ExcelPageCountSize1 { get; set; }
        private static Application ExcelPageCountSize2 { get; set; }

        //Task status indicator
        public bool Task1Runing { get; private set; }

        public bool Task2Runing { get; private set; }

        //Two queue used to store tasks
        public ConcurrentQueue<Task> PageSizeCountTaskQueue1 { get; private set; } = new ConcurrentQueue<Task>();

        public ConcurrentQueue<Task> PageSizeCountTaskQueue2 { get; private set; } = new ConcurrentQueue<Task>();
        public CancellationToken CancellationToken { get; private set; }

        //Update status function injected from outside
        public Action<string> UpdateStatus { get; set; }
        #endregion

        #region Constructors
        public WorkbookPageSizeCountExtractor(CancellationToken cancellationToken)
        {
            CancellationToken = cancellationToken;
        }

        ~WorkbookPageSizeCountExtractor()
        {
            if (ExcelPageCountSize1 != null)
            {
                //ExcelPageCountSize1.Workbooks.Close();
                try
                {
                    ExcelPageCountSize1.Quit();
                    Marshal.ReleaseComObject(ExcelPageCountSize1);
                }
                catch { }
                ExcelPageCountSize1 = null;
            }
            if (ExcelPageCountSize2 != null)
            {
                //ExcelPageCountSize2.Workbooks.Close();
                try
                {
                    ExcelPageCountSize2.Quit();
                    Marshal.ReleaseComObject(ExcelPageCountSize2);
                }
                catch { }
                ExcelPageCountSize2 = null;
            }
        }
        #endregion

        #region Public methods
        //Used to set a new cancel token from outside
        public void SetCancelToken(CancellationToken cancellationToken)
        {
            CancellationToken = cancellationToken;
        }

        //public glue method
        public async void GetPageCountSize(IEnumerable<WorkbookInfo> workbookInfos)
        {
            await StartExcelPageCountSize();
            EnqueuePageCountSizeTask(workbookInfos);
            ExecuteQueue();
        }
        #endregion

        #region Private methods
        //Start two excels
        private async Task StartExcelPageCountSize()
        {
            if (ExcelPageCountSize1 == null)
            {
                await Task.Run(() =>
                {
                    ExcelPageCountSize1 = new Application();
                    ExcelPageCountSize1.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                });
            }
            if (ExcelPageCountSize2 == null)
            {
                await Task.Run(() =>
                {
                    ExcelPageCountSize2 = new Application();
                    ExcelPageCountSize2.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                });
            }
        }

        //Split workboolinfo list into two parts and generate two tasks into the queue
        private void EnqueuePageCountSizeTask(IEnumerable<WorkbookInfo> workbookInfos)
        {
            if (workbookInfos.Count() > 1)
            {
                var workbookSplits = workbookInfos.Split(2);
                var task1 = new Task(() =>
                {
                    var t = GetPageCountAndSize(ExcelPageCountSize1, workbookSplits.ElementAt(0));
                    t.Wait();
                });
                var task2 = new Task(() =>
                {
                    var t = GetPageCountAndSize(ExcelPageCountSize2, workbookSplits.ElementAt(1));
                    t.Wait();
                });
                PageSizeCountTaskQueue1.Enqueue(task1);
                PageSizeCountTaskQueue2.Enqueue(task2);
            }
            else
            {
                var task1 = new Task(() =>
                {
                    var t = GetPageCountAndSize(ExcelPageCountSize1, workbookInfos);
                    t.Wait();
                });
                PageSizeCountTaskQueue1.Enqueue(task1);
            }
        }

        //Execute queue
        private void ExecuteQueue()
        {
            if (!Task1Runing)
                Task.Run(() =>
                {
                    while (PageSizeCountTaskQueue1.Count > 0)
                    {
                        Task1Runing = true;
                        Task t;
                        if (PageSizeCountTaskQueue1.TryDequeue(out t))
                        {
                            t.Start();
                            t.Wait();
                        }
                    }
                    Task1Runing = false;
                });

            if (!Task2Runing)
                Task.Run(() =>
                {
                    while (PageSizeCountTaskQueue2.Count > 0)
                    {
                        Task2Runing = true;
                        Task t;
                        if (PageSizeCountTaskQueue2.TryDequeue(out t))
                        {
                            t.Start();
                            t.Wait();
                        }
                    }
                    Task2Runing = false;
                });
        }

        //Batch method
        private async Task GetPageCountAndSize(Application excel, IEnumerable<WorkbookInfo> workbookInfos)
        {
            foreach (var workbookInfo in workbookInfos)
            {
                if (workbookInfo.IsWorksheetPageCountSizeObtained) continue;
                workbookInfo.ShowProgressBar = true;
                await GetPageCountAndSize(excel, workbookInfo);
                workbookInfo.ShowProgressBar = false;
                if (CancellationToken.IsCancellationRequested)
                    break;
            }
        }

        //Genuine method
        private async Task GetPageCountAndSize(Application excel, WorkbookInfo workbookInfo)
        {
            var result = await workbookInfo.GetWorksheetPageCountAndSize(excel, CancellationToken, UpdateStatus);
            if (result)
                workbookInfo.IsWorksheetPageCountSizeObtained = true;
        }
        #endregion
    }
}