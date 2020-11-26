namespace ExcelToPaper.DataModels
{
    public class PrintSettings
    {
        //Is print to paper.
        public bool PrintToPaper { get; set; }
        public bool PrintToPdf { get; set; } = true;

        //Export to a separate folders
        public bool ExportToSeparateFolder { get; set; } = true;
        public bool ExportToSingleFolder { get; set; } = false;

        //Attach workbook name before work sheet name
        public bool AttachWorkbookNameBeforeWorksheet { get; set; }

        //The single folder path that pdf will be printed to.
        public string SingleFolderPath { get; set; }
        //Don't merge
        public bool MergeNothing { get; set; } = true;
        //Merge to file separately
        public bool MergeToFileSeparately { get; set; }
        //Merge to all to a single file
        public bool MergeToSingleFile { get; set; }
        public bool MergeDeleteOriginFile { get; set; } = false;
    }
}