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
    /// <item>Microsoft Access Database Engine 2010 Redistributable: https://www.microsoft.com/en-us/download/details.aspx?id=13255</item>
    /// </list>
    /// </remarks>
    public class ExcelProvider : IExcelProvider
    {
        private string sheetName;
        private string fileName;
        private FileType fileType;

        private readonly List<IExcelRow> rows = new List<IExcelRow>();
        private readonly List<IExcelColumn> columns = new List<IExcelColumn>();

        /// <summary>
        /// File name of Excel XLSX/XLS or CSV file.
        /// </summary>
        public string FileName
        {
            get { return fileName; }
        }

        /// <summary>
        /// Type of File: XLSX, XLS or CSV.
        /// </summary>
        public FileType FileType
        {
           get { return fileType; }
        }

        /// <summary>
        /// Name of worksheet.
        /// </summary>
        public string SheetName
        {
            get { return sheetName; }
        }

        /// <summary>
        /// Collection of Excel rows.
        /// </summary>
        public IEnumerable<IExcelRow> Rows
        {
            get { return rows; }
        }

        /// <summary>
        /// Collection of definitions of Excel columns.
        /// </summary>

        public IEnumerable<IExcelColumn> Columns
        {
            get { return columns; }
        }

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
            this.fileName = fileName;
            this.fileType = GetFileType();
            this.sheetName = sheetName;

            if (!File.Exists(fileName))
            {
                throw new FileNotFoundException($"File {fileName} does not exist");
            }

            if ((fileType != FileType.Csv) && (string.IsNullOrEmpty(sheetName)))
            {
                throw new ArgumentNullException(nameof(sheetName), $"Worksheet name is required for file {fileName}");
            }

            LoadWorksheet();
        }

        private FileType GetFileType()
        {
            var extension = Path.GetExtension(fileName).ToUpper();
            if (extension == ".XLSX")
            {
                return FileType.Xlsx;
            }
            else if (extension == ".XLS")
            {
                return FileType.Xls;
            }
            else if (extension == ".CSV")
            {
                return FileType.Csv;
            }
            else
            {
                throw new ArgumentException($"File {fileName} with extension {extension} is not supported");
            }
        }

        private string GetConnectionString()
        {
            switch (fileType)
            {
                case FileType.Xls:
                    return $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={fileName};Extended Properties=""Excel 8.0;HDR=YES;""";
                case FileType.Csv:
                    // https://msdn.microsoft.com/en-us/library/ms974559.aspx
                    // The delimiter of the CSV can be specified in the registry at the following location: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text
                    // Format can be "TabDelimited", "CSVDelimited" or "Delimited(;)"
                    // Or create a schema.ini file in the same folder as the CSV file where you specify the delimiter
                    // "HDR=Yes;" indicates that the first row contains column names, not data. "HDR=No;" indicates the opposite.
                    return $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={Path.GetDirectoryName(fileName)};Extended Properties=""text;HDR=Yes;FMT=Delimited;""";
                default:
                    return $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fileName};Extended Properties=""Excel 12.0 Xml;HDR=YES""";
            }
        }

        private string GetCommandText()
        {
            switch (fileType)
            {
                case FileType.Csv:
                    return $"SELECT * FROM {Path.GetFileName(fileName)}";
                default:
                    return $"SELECT * FROM [{sheetName}$]";
            }
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
                                columns.Add(new ExcelColumn(i, reader.GetName(i), reader.GetFieldType(i)));
                            }

                            int rowCount = 1;
                            // Run through records and create rows with cells with contain the values
                            while (reader.Read())
                            {
                                var newRow = new ExcelRow(rowCount++, columns);
                                for (int index = 0; index < reader.FieldCount; index++)
                                {
                                    newRow.AddCell(reader[index]);
                                }
                                rows.Add(newRow);
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
