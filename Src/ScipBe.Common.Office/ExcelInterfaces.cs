// ==============================================================================================
// Namespace   : ScipBe.Common.Office
// Author      : Stefan Cruysberghs
// Website     : http://www.scip.be
// Status      : Open source - MIT License
// ==============================================================================================

using System;

namespace ScipBe.Common.Office
{
  /// <summary>
  /// Excel version
  /// </summary>
  /// <remarks>
  /// Connection strings
  /// <list type="bullet">
  /// <item>Excel 2007 : @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=FileName;Extended Properties=""Excel 12.0 Xml;HDR=YES"""</item>
  /// <item>Excel 2000-2003 : @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Path;Extended Properties=""text;HDR=Yes;FMT=Delimited;"""</item>
  /// <item>CSV : @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=FileName;Extended Properties=""Excel 8.0;HDR=YES;"""</item>
  /// </list>  
  /// </remarks>
  public enum ExcelVersion
  {
    /// <summary>
    /// Excel 2000 (v9, XLS)
    /// </summary>
    Excel2000,
    /// <summary>
    /// Excel XP (v10, XLS)
    /// </summary>
    ExcelXp,
    /// <summary>
    /// Excel 2003 (v11, XLS)
    /// </summary>
    Excel2003,
    /// <summary>
    /// Excel 2007 (v12, XLS and XLSX) 
    /// </summary>
    Excel2007,
    /// <summary>
    /// CSV, comma separated ASCII file
    /// </summary>
    Csv
  };
  
  /// <summary>
  /// Interface for column definition of Excel worksheet
  /// </summary>
  public interface IExcelColumn
  {
    /// <summary>
    /// Column index, 1 - 999999
    /// </summary>
    int Index { get; }

    /// <summary>
    /// Column header, A - Z, AA - ZZ, ...
    /// </summary>
    string Header { get; }

    /// <summary>
    /// Column name. String name in first row of Excel worksheet
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Column type (string, int, datetime, ...)
    /// </summary>
    Type Type { get; }
  }

  /// <summary>
  /// Interface for row with cell values of Excel worksheet
  /// </summary>
  /// <example><code>
  /// var excel = ExcelProvider.Create(@"Path\Persons.xlsx", "Persons");
  /// var query = 
  ///   from r in excel.Rows
  ///   select new
  /// 	{
  /// 	  ID =  r[1],
  /// 	  ID2 = r["A"],
  /// 	  ID3 = r.Get&lt;int&gt;(1),
  /// 	  ID4 = r.Get&lt;string&gt;(1),
  /// 	  ID5 = r.Get&lt;int&gt;("A"),
  /// 	  ID6 = r.Get&lt;string&gt;("A"),
  /// 	  ID7 = r.GetByName&lt;int&gt;("ID"),
  /// 	  ID8 = r.GetByName&lt;string&gt;("ID"),
  /// 	  FirstName = r[2],
  /// 	  LastName = r.Get&lt;string&gt;(3),
  /// 	  Country = r.GetByName&lt;string&gt;("Country"),
  /// 	  BirthDate = r.Get&lt;DateTime&gt;(5),
  /// 	  BirthDate2 = r[5].ToString(),
  /// 	  BirthDate3 = r["E"],
  /// 	  BirthDate4 = r.GetByName&lt;string&gt;("BirthDate"),
  /// 	  BirthDate5 = r.Get&lt;DateTime&gt;("E"),
  /// 	  BirthDate6 = r.GetByName&lt;DateTime&gt;("BirthDate").AddMonths(1),
  /// 	  Row = r.ToString()
  /// 	};
  /// </code></example>
  public interface IExcelRow
  {
    /// <summary>
    /// Row index, 1 - 999999
    /// </summary>
    int Index { get; }

    /// <summary>
    /// Get cell value as Object. Column index is given.
    /// </summary>
    /// <param name="columnIndex">Column index, starts with 1</param>
    /// <returns>Value of cell as Object</returns>
    Object this[int columnIndex] { get; }

    /// <summary>
    /// Get cell value as Object. Column header is given.
    /// </summary>
    /// <param name="columnHeader">Column header, starts with A</param>
    /// <returns>Value of cell as Object</returns>
    Object this[string columnHeader] { get; }

    /// <summary>
    /// Get cell value as given class type. Column index is given.
    /// </summary>
    /// <typeparam name="T">Class type of cell value</typeparam>
    /// <param name="columnIndex">Column index, starts with 1</param>
    /// <returns>Value of cell</returns>
    T Get<T>(int columnIndex);

    /// <summary>
    /// Get cell value as given class type. Column header is given.
    /// </summary>
    /// <typeparam name="T">Class type of cell value</typeparam>
    /// <param name="columnHeader">Column headers, starts with A</param>
    /// <returns>Value of cell</returns>
    T Get<T>(string columnHeader);

    /// <summary>
    /// Get cell value as given class type. Column name (from first row) is given.
    /// </summary>
    /// <typeparam name="T">Class type of cell value</typeparam>
    /// <param name="columnName">Column name, is string name in first row of Excel worksheet</param>
    /// <returns>Value of cell</returns>
    T GetByName<T>(string columnName);

    /// <summary>
    /// Get string representation of row
    /// </summary>
    /// <returns>Comma seperated list of all cell values</returns>
    string ToString();
  }
}