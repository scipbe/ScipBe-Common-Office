using System.Collections.Generic;

namespace ScipBe.Common.Office.Excel
{
    public interface IExcelProvider
    {
        /// <summary>
        /// File name of Excel XLSX/XLS or CSV file.
        /// </summary>
        string FileName { get; }

        /// <summary>
        /// Type of File: XLSX, XLS or CSV.
        /// </summary>
        FileType FileType { get; }

        /// <summary>
        /// Name of worksheet.
        /// </summary>
        string SheetName { get; }

        /// <summary>
        /// Collection of definitions of Excel columns.
        /// </summary>
        List<IExcelColumn> Columns { get; }

        /// <summary>
        /// Collection of Excel rows.
        /// </summary>
        List<IExcelRow> Rows { get; }

        /// <summary>
        /// Load XLSX, XLS or CSV file of given worksheet.
        /// </summary>
        /// <param name="fileName">Name of XLSX, XLS or CSV file.</param>
        /// <param name="sheetName">Name of worksheet. Required for XLS or XLSX file. Can be empty for CSV file.</param>
        /// <remarks>
        /// The file name of the CSV file should not contains spaces.
        /// The first row of CSV file needs a to contain the column names.
        /// The delimiter of the CSV can be specified in the registry at the following location: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text.
        /// Format can be "TabDelimited", "CSVDelimited" or "Delimited(;)".
        /// Or create a schema.ini file in the same folder as the CSV file where you specify the delimiter.
        /// </remarks>
        void Load(string fileName, string sheetName = null);
    }
}