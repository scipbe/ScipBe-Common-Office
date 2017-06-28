using System;

namespace ScipBe.Common.Office.Excel
{
    /// <summary>
    /// Interface for column definition of Excel worksheet.
    /// </summary>
    public interface IExcelColumn
    {
        /// <summary>
        /// Column index, 1 - 999999.
        /// </summary>
        int Index { get; }

        /// <summary>
        /// Column header, A - Z, AA - ZZ, ...
        /// </summary>
        string Header { get; }

        /// <summary>
        /// Column name. String name in first row of Excel worksheet.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Column type (string, int, datetime, ...).
        /// </summary>
        Type Type { get; }
    }
}