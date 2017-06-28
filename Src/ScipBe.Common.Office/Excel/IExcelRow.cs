using System;

namespace ScipBe.Common.Office.Excel
{
    /// <summary>
    /// Interface for row with cell values of Excel worksheet.
    /// </summary>
    public interface IExcelRow
    {
        /// <summary>
        /// Row index, 1 - 999999.
        /// </summary>
        int Index { get; }

        /// <summary>
        /// Get cell value as Object. Column index is given.
        /// </summary>
        /// <param name="columnIndex">Column index, starts with 1.</param>
        /// <returns>Value of cell as Object.</returns>
        Object this[int columnIndex] { get; }

        /// <summary>
        /// Get cell value as Object. Column header is given.
        /// </summary>
        /// <param name="columnHeader">Column header, starts with A.</param>
        /// <returns>Value of cell as Object.</returns>
        Object this[string columnHeader] { get; }

        /// <summary>
        /// Get cell value as given class type. Column index is given.
        /// </summary>
        /// <typeparam name="T">Class type of cell value</typeparam>
        /// <param name="columnIndex">Column index, starts with 1.</param>
        /// <returns>Value of cell</returns>
        T Get<T>(int columnIndex);

        /// <summary>
        /// Get cell value as given class type. Column header is given.
        /// </summary>
        /// <typeparam name="T">Class type of cell value.</typeparam>
        /// <param name="columnHeader">Column headers, starts with A.</param>
        /// <returns>Value of cell</returns>
        T Get<T>(string columnHeader);

        /// <summary>
        /// Get cell value as given class type. Column name (from first row) is given.
        /// </summary>
        /// <typeparam name="T">Class type of cell value.</typeparam>
        /// <param name="columnName">Column name, is string name in first row of Excel worksheet.</param>
        /// <returns>Value of cell</returns>
        T GetByName<T>(string columnName);

        /// <summary>
        /// Get string representation of row.
        /// </summary>
        /// <returns>Comma seperated list of all cell values.</returns>
        string ToString();
    }
}