namespace ScipBe.Common.Office.Excel
{
    /// <summary>
    /// File type.
    /// </summary>
    /// <remarks>
    /// Connection strings
    /// <list type="bullet">
    /// <item>XLSX: @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=FileName;Extended Properties=""Excel 12.0 Xml;HDR=YES"""</item>
    /// <item>XLS: @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=FileName;Extended Properties=""Excel 8.0;HDR=YES"""</item>
    /// </list>  
    /// </remarks>
    public enum FileType
  {
    /// <summary>
    /// Excel 97-2003 (v8-v11).
    /// </summary>
    Xls,
    /// <summary>
    /// Excel 2007-2019 (v12-v16). 
    /// </summary>
    Xlsx
  };
}