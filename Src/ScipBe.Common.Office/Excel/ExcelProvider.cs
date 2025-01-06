using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;

namespace ScipBe.Common.Office.Excel
{
    /// <summary>
    /// Excel Provider (LINQ to Excel).
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>Author: Stefan Cruysberghs</item>
    /// <item>Website: http://www.scip.be</item>
    /// <item>Article: Querying Excel worksheets with LINQ : http://www.scip.be/index.php?Page=ArticlesNET25</item>
    /// <item>Microsoft Access Database Engine 2016 Redistributable: https://www.microsoft.com/en-us/download/details.aspx?id=54920</item>
    /// </list>
    /// </remarks>
    public class ExcelProvider : IExcelProvider
    {
        /// <summary>
        /// File name of Excel XLSX or XLS  file.
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// Type of File: XLSX or XLS.
        /// </summary>
        public FileType FileType { get; private set; }

        /// <summary>
        /// Name of worksheet.
        /// </summary>
        public string SheetName { get; private set; }

        /// <summary>
        /// Collection of Excel rows.
        /// </summary>
        public List<IExcelRow> Rows { get; private set; } = new List<IExcelRow>();

        /// <summary>
        /// Collection of definitions of Excel columns.
        /// </summary>

        public List<IExcelColumn> Columns { get; private set; } = new List<IExcelColumn>();

        /// <summary>
        /// Constructor of ExcelProvider, it will load Excel worksheet or CSV into Rows and Columns collections.
        /// </summary>
        /// <param name="fileName">Name of XLSX, XLS or CSV file.</param>
        /// <param name="sheetName">Name of worksheet. Required for XLS or XLSX file. Can be empty for CSV file.</param>
        /// <remarks>
        /// The first row of CSV file needs a to contain the column names.
        /// The delimiter of the CSV can be specified in the registry at the following location: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text.
        /// Format can be "TabDelimited", "CSVDelimited" or "Delimited(;)".
        /// Or create a schema.ini file in the same folder as the CSV file where you specify the delimiter.
        /// </remarks> 
        public ExcelProvider(string fileName, string sheetName = null)
        {
            Load(fileName, sheetName);
        }

        /// <summary>
        /// Constructor of ExcelProvider. Call Load afterwards to load worksheet data.
        /// </summary>
        public ExcelProvider()
        {
        }

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
        public void Load(string fileName, string sheetName = null)
        {
            this.FileName = fileName;
            this.FileType = GetFileType();
            this.SheetName = sheetName;

            if (!File.Exists(fileName))
            {
                throw new FileNotFoundException($"File {fileName} does not exist");
            }

            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName), $"Worksheet name is required for file {fileName}");
            }

            LoadWorksheet();
        }

        private FileType GetFileType()
        {
            var extension = Path.GetExtension(FileName).ToUpper();
            switch (extension)
            {
                case ".XLSX":
                    return FileType.Xlsx;
                case ".XLS":
                    return FileType.Xls;
                default:
                    throw new ArgumentException($"File {FileName} with extension {extension} is not supported");
            }
        }

        private string GetConnectionString()
        {
            switch (FileType)
            {
                case FileType.Xls:
                    return $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={FileName};Extended Properties=""Excel 8.0;HDR=YES""";
                default:
                    return $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={FileName};Extended Properties=""Excel 12.0 Xml;HDR=YES""";
            }
        }

        private string GetCommandText()
        {
            return $"SELECT * FROM [{SheetName}$]";
        }

        private void LoadWorksheet()
        {
            string connectionString = GetConnectionString();

            using (var connection = new OleDbConnection(connectionString))
            {
                // Get OleDB connection to Excel (XLS/XLSX) or CSV file and open it with a DataReader.
                try
                {
                    connection.Open();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = GetCommandText();
                        using (var reader = command.ExecuteReader())
                        {
                            // Run through fields and create column objects
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Columns.Add(new ExcelColumn(i, reader.GetName(i), reader.GetFieldType(i)));
                            }

                            int rowCount = 1;
                            // Run through records and create rows with cells with contain the values
                            while (reader.Read())
                            {
                                var newRow = new ExcelRow(rowCount++, Columns);
                                for (int index = 0; index < reader.FieldCount; index++)
                                {
                                    newRow.AddCell(reader[index]);
                                }
                                Rows.Add(newRow);
                            }
                        }
                    }
                }
                finally
                {
                    connection.Close();
                }
            }
        }
    }
}
