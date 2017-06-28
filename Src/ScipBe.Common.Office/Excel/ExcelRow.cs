using System;
using System.Collections.Generic;

namespace ScipBe.Common.Office.Excel
{
    internal class ExcelRow : IExcelRow
    {
        private readonly List<IExcelColumn> columns;
        private readonly List<Object> cells = new List<Object>();

        public ExcelRow(int index, List<IExcelColumn> columns)
        {
            Index = index;
            this.columns = columns;
        }

        internal void AddCell(Object data)
        {
            cells.Add(data);
        }

        private int CheckColumn(int columnIndex)
        {
            if ((columnIndex > 0) && (columnIndex <= columns.Count))
                return columnIndex - 1;
            return -1;
        }

        private int CheckColumn(string columnHeader)
        {
            int columnIndex = columns.FindIndex(c => c.Header == columnHeader);
            return columnIndex;
        }

        private int CheckColumnByName(string columnName)
        {
            int columnIndex = columns.FindIndex(c => c.Name == columnName);
            return columnIndex;
        }

        private static T ConvertTo<T>(object data)
        {
            try
            {
                if ((data is DBNull) || (data == null))
                {
                    return default(T);
                }
                return (T)Convert.ChangeType(data, typeof(T));
            }
            catch
            {
                return default(T);
            }
        }

        public int Index { get; internal set; }

        public Object this[int columnIndex]
        {
            get
            {
                if (CheckColumn(columnIndex) > -1)
                {
                    return cells[columnIndex - 1];
                }
                return null;
            }
        }

        public Object this[string columnHeader]
        {
            get
            {
                int columnIndex = CheckColumn(columnHeader);
                if (columnIndex > -1)
                {
                    return cells[columnIndex];
                }
                return null;
            }
        }

        public T Get<T>(int columnIndex)
        {
            if (CheckColumn(columnIndex) > -1)
            {
                return ConvertTo<T>(cells[columnIndex - 1]);
            }
            return default(T);
        }

        public T Get<T>(string columnHeader)
        {
            int columnIndex = CheckColumn(columnHeader);
            if (columnIndex > -1)
            {
                return ConvertTo<T>(cells[columnIndex]);
            }
            return default(T);
        }

        public T GetByName<T>(string columnName)
        {
            int columnIndex = CheckColumnByName(columnName);
            if (columnIndex > -1)
            {
                return ConvertTo<T>(cells[columnIndex]);
            }
            return default(T);
        }

        public new string ToString()
        {
            string values = "";
            foreach (var cell in cells)
            {
                values += ";" + cell.ToString();
            }
            return values.Substring(1);
        }
    }
}