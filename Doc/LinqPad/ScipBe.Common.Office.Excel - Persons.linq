<Query Kind="Program">
  <NuGetReference>ScipBe.Common.Office.Excel</NuGetReference>
  <Namespace>ScipBe.Common.Office.Excel</Namespace>
</Query>

void Main()
{
    Util.RawHtml($"<h2>ScipBe.Common.Office.Excel - LINQ to Excel - Load Excel worksheet or CSV file</h2>").Dump();

    var excel = new ExcelProvider($@"{Path.GetDirectoryName(Util.CurrentQueryPath)}\Persons.xlsx", "Persons");
    //var excel = new ExcelProvider($@"{Path.GetDirectoryName(Util.CurrentQueryPath)}\Persons.xls", "Persons");
    //var excel = new ExcelProvider($@"{Path.GetDirectoryName(Util.CurrentQueryPath)}\PersonsTab.csv");
    //var excel = new ExcelProvider($@"{Path.GetDirectoryName(Util.CurrentQueryPath)}\PersonsComma.csv");
    //var excel = new ExcelProvider($@"{Path.GetDirectoryName(Util.CurrentQueryPath)}\PersonsSemicolumn.csv");

	excel.Columns.Dump("Columns");

	var query =
	from r in excel.Rows
	select new
	{
		ID = r[1],
		ID2 = r["A"],
		ID3 = r.Get<int>(1),
		ID4 = r.Get<string>(1),
		ID5 = r.Get<int>("A"),
		ID6 = r.Get<string>("A"),
		ID7 = r.GetByName<int>("ID"),
		ID8 = r.GetByName<string>("ID"),
		FirstName = r[2],
		LastName = r.Get<string>(3),
		Country = r.GetByName<string>("Country"),
		BirthDate = r.Get<DateTime>(5),
		BirthDate2 = r[5].ToString(),
		BirthDate3 = r["E"],
		BirthDate4 = r.GetByName<string>("BirthDate"),
		BirthDate5 = r.Get<DateTime>("E"),
		BirthDate6 = r.GetByName<DateTime>("BirthDate").AddMonths(1),
		Row = r.ToString()
	};

	query.Dump("Rows");
}