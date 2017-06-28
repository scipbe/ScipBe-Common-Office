using System;

namespace ScipBe.Common.Office.Excel
{
    internal class ExcelColumn : IExcelColumn
    {
        public ExcelColumn(int index, string name, Type type)
        {
            Index = index + 1;
            Name = name;
            Type = type;
            Header = CalcHeader(index + 1);
        }

        private static string CalcHeader(int index)
        {
            int number = index;
            string headerText = "";
            while (number > 0)
            {
                headerText = (Char)(((number - 1) % 26) + 65) + headerText;
                number = (number - 1) / 26;
            }
            return headerText;
        }

        public int Index { get; set; }
        public string Header { get; set; }
        public string Name { get; set; }
        public Type Type { get; set; }
    }
}