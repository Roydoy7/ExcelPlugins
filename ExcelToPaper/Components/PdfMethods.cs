using ExcelToPaper.DataModels;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelToPaper.Components
{
    public static class PdfMethods
    {
        public static void MergePdf(PrintResult printResult, Action<string> updateStatus = null)
        {
            var workbookNameNoExt = Path.GetFileNameWithoutExtension(printResult.WorkbookPath);
            var folderPath = printResult.OutputFolderPath;
            var mergedFilePath = folderPath + "\\" + workbookNameNoExt + ".pdf";
            updateStatus?.Invoke("マージ中... " + mergedFilePath);
            MergePdf(printResult.PrintedPdfPaths, mergedFilePath, updateStatus);
            updateStatus?.Invoke("マージ中... 完成");
        }

        public static void MergePdf(IEnumerable<string> filePaths, string mergedFilePath, Action<string> updateStatus = null)
        {
            if (!filePaths.Any()) return;
            if (filePaths.Count() == 1) return;

            var pdfDocOut = new PdfDocument();

            foreach (var filePath in filePaths)
            {
                try
                {
                    var pdfDocIn = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
                    foreach (var page in pdfDocIn.Pages)
                        pdfDocOut.AddPage(page);
                }
                catch(Exception e) 
                {
                    updateStatus?.Invoke(e.Message);
                }
            }

            pdfDocOut.Save(mergedFilePath);
        }

        public static void DeletePdf(IEnumerable<string> filePaths)
        {
            if (!filePaths.Any()) return;
            foreach (var filePath in filePaths)
                File.Delete(filePath);
        }
    }
}
