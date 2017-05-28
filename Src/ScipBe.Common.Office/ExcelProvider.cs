// ==============================================================================================
// Namespace   : ScipBe.Common.Office
// Class(es)   : ExcelProvider (LINQ to Excel)
// Version     : 1.3
// Author      : Stefan Cruysberghs
// Website     : http://www.scip.be
// Date        : August 2008 - December 2008
// Status      : Open source - MIT License
// ==============================================================================================

using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;

namespace ScipBe.Common.Office
{
  /// <summary>
  /// Excel Provider (LINQ to Excel)
  /// </summary>
  /// <remarks>
  /// <list type="bullet">
  /// <item>Author: Stefan Cruysberghs</item>
  /// <item>Website: http://www.scip.be</item>
  /// <item>Article: Querying Excel worksheets with LINQ : http://www.scip.be/index.php?Page=ArticlesNET25</item>
  /// <item>Source code: http://www.scip.be/index.php?Page=ComponentsNETOfficeItems</item>
  /// </list>
  /// </remarks>
  /// <example><code>
  /// // LINQPad example
  /// var excel = ExcelProvider.Create(@"Path\Persons.xls", "Persons");
  /// //var excel = ExcelProvider.Create(@"Path\Persons.xls", "Persons", ExcelVersion.Excel2007);
  /// //var excel = ExcelProvider.Create(@"Path\Persons.xl", "Persons", ExcelVersion.ExcelXp);
  /// //var excel = ExcelProvider.Create(@"Path\Persons.csv", "", ExcelVersion.Csv);
  /// 
  /// excel.Columns.Dump("ExcelProvider : columns");
  /// 
  /// var query = from r in excel.Rows
  ///             select new
  /// 			      {
  /// 			        ID =  r[1],
  /// 			        FirstName = r.Get&lt;string&gt;(2),
  /// 			        LastName = r.Get&lt;string&gt;("C"),
  /// 			        Country = r.GetByName&lt;string&gt;("Country"),
  /// 			        BirthDate = r.GetByName&lt;DateTime&gt;("BirthDate")
  /// 			      };
  /// 
  /// query.Dump("ExcelProvider : cell values from rows");
  /// </code></example>
  public class ExcelProvider
  {
    private string sheetName;
    private string fileName;
    private ExcelVersion version;

    private readonly List<IExcelRow> rows = new List<IExcelRow>();
    private readonly List<IExcelColumn> columns = new List<IExcelColumn>();

    /// <summary>
    /// File name of Excel workbook or CSV file
    /// </summary>
    public string FileName
    {
      get { return fileName; }
    }

    /// <summary>
    /// Name of worksheet
    /// </summary>
    public string SheetName
    {
      get { return sheetName; }
    }

    /// <summary>
    /// Collection of Excel rows
    /// </summary>
    public IEnumerable<IExcelRow> Rows
    {
      get { return rows; }
    }
    
    /// <summary>
    /// Collection of definitions of Excel columns
    /// </summary>
    public IEnumerable<IExcelColumn> Columns
    {
      get { return columns; }
    }

    /// <summary>
    /// Factory method to load Excel worksheet into Rows and Columns collections.
    /// </summary>
    /// <param name="fileName">Name of XLS, XLSX or CSV file</param>
    /// <param name="sheetName">Name of worksheet</param>
    /// <param name="version">Excel version</param>
    /// <returns>Instance of ExcelProvider</returns>
    public static ExcelProvider Create(string fileName, string sheetName, ExcelVersion version)
    {
      ExcelProvider excelProvider = new ExcelProvider()
      {
        sheetName = sheetName,
        fileName = fileName,
        version = version
      };

      if (!File.Exists(fileName))
      {
        throw new FileNotFoundException("File does not exists");
      }

      if ((version != ExcelVersion.Csv) && (string.IsNullOrEmpty(sheetName)))
      {
        throw new ArgumentNullException("sheetName", "Worksheet name is required");
      }

      excelProvider.Load();
      return excelProvider;
    }

    /// <summary>
    /// Factory method to load Excel worksheet into Rows and Columns collections.
    /// </summary>
    /// <param name="filePath">Name of XLS or XLSX file</param>
    /// <param name="sheetName">Name of worksheet</param>
    /// <returns>Instance of ExcelProvider</returns>
    public static ExcelProvider Create(string filePath, string sheetName)
    {
      return Create(filePath, sheetName, ExcelVersion.Excel2007);
    }

    private string GetConnectionString()
    {
      switch (version)
      {
        case ExcelVersion.Excel2007:
          return string.Format(
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""",
            fileName);
        case ExcelVersion.Csv:
          return string.Format(
            @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=Yes;FMT=Delimited;""",
            Path.GetDirectoryName(fileName));
        default:
          return string.Format(
            @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;""",
            fileName);
      }
    }

    private string GetCommandText()
    {
      switch (version)
      {
        case ExcelVersion.Csv:
          return string.Format("SELECT * FROM {0}", Path.GetFileName(fileName));
        default:
          return string.Format("SELECT * FROM [{0}$]", sheetName);
      }
    }

    private void Load()
    {
      string connectionString = GetConnectionString();

      using (OleDbConnection connection = new OleDbConnection(connectionString))
      {
        // Get OleDB connection to Excel (XLS/XLSX) or CSV file
        // and open it with a DataReader.
        connection.Open();
        using (OleDbCommand command = connection.CreateCommand())
        {
          command.CommandText = GetCommandText();
          using (OleDbDataReader reader = command.ExecuteReader())
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
              ExcelRow newRow = new ExcelRow(rowCount++, columns);
              for (int index = 0; index < reader.FieldCount; index++)
              {
                newRow.AddCell(reader[index]);
              }
              rows.Add(newRow);
            }
          }
        }
      }
    }
  }
}
