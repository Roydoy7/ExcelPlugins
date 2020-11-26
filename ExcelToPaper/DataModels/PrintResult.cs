using System.Collections.Generic;

namespace ExcelToPaper.DataModels
{
    public class PrintResult
    {
        //Workbook file path
        public string WorkbookPath { get; set; }
        //Output folder path
        public string OutputFolderPath { get; set; }
        //Printed pdf file paths
        public List<string> PrintedPdfPaths { get; private set; } = new List<string>();
    }
}
