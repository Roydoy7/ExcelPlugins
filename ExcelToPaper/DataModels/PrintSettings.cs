namespace ExcelToPaper.DataModels
{
    internal class PrintSettings
    {
        //Is print to paper.
        public bool PrintToPaper { get; set; } = false;
        
        //Export to a single folder
        public bool ExportToSingleFolder { get; set; } = false;

        //Attach workbook name before work sheet name
        public bool AttachWorkbookNameBeforeWorksheet { get; set; }

        //The single folder path that pdf will be printed to.
        public string SingleFolderPath { get; set; }
    }
}