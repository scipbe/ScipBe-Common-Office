![Logo](Doc/Images/ScipBe.Common.Office.png) 
# ScipBe-Common-Office
### Linq to Excel, Outlook and OneNote

The ScipBe.Common.Office namespace contains 3 classes: ExcelProvider (LINQ to Excel), OutlookProvider (LINQ to Outlook) and OneNoteProvider (LINQ to OneNote). 
- The ExcelProvider loads an Excel worksheet or CSV file and provides column definition and row collections. 
- The OutlookProvider is a wrapper class which provides collections to data of Outlook (AppointmentItems, ContactItems, MailItems, TaskItems, ...). 
- The OneNoteProvider provides collections of Notebooks, Sections and Pages by parsing the XML hierarchy tree of OneNote. 
- All collections are IEnumerable so you can query them with LINQ. 
- There are also 3 separated projects with only Excel, Outlook and OneNote provider.

Examples
=================================================================

- See scripts in Doc\LinqPad folder
- See class diagrams in Doc\Diagrams folder or in solution

Links
=================================================================

- [Homepage](http://www.scip.be)
- [Documentation and examples](http://www.scip.be/index.php?Page=ComponentsNETOfficeItems)
- [Author Stefan Cruysberghs](http://www.scip.be/index.php?Page=AboutMe)
- [GitHub repository](https://github.com/scipbe/ScipBe-Common-Office)
- [NuGet package Office](https://www.nuget.org/packages/ScipBe.Common.Office)
- [NuGet package Excel](https://www.nuget.org/packages/ScipBe.Common.Office.Excel)
- [NuGet package Outlook](https://www.nuget.org/packages/ScipBe.Common.Office.Outlook)
- [NuGet package OneNote](https://www.nuget.org/packages/ScipBe.Common.Office.OneNote)

Remarks
=================================================================

- ExcelProvider
  - The ExcelProvider supports XLSX (Excel 2007-2016, v12-v16), XLS (Excel 97-2003, v8-v11) and CSV (comma, semicolumn or tab delimited ASCII file) files but it requires the installation of the Microsoft Access Database Engine 2010 Redistributable: https://www.microsoft.com/en-us/download/details.aspx?id=13255
  - CSV files:
    - The file name of the CSV file should not contains spaces
	- The first row of CSV file needs a to contain the column names
    - The delimiter of the CSV can be specified in the registry at the following location: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text
    - Format can be "TabDelimited", "CSVDelimited" or "Delimited(;)"
    - Or create a schema.ini file in the same folder as the CSV file where you specify the delimiter
	- See https://msdn.microsoft.com/en-us/library/ms974559.aspx for more information

- OutlookProvider
  - The OutlookProvider exposes collections of classes of the Microsoft.Office.Interop.Outlook assembly so you have full access to all properties and it also supports adding, updating and deleting data in Outlook.

History
=================================================================
- Version 3.0.0 (May 2022)
  - Migrated to .NET Standard 2.0 to allow .NET Core projects to consume (#8)
  - Switched OneNoteProvider to be static, and to release the COM interop library around every call, and adding a GetContent() method to PageItem (#11, #3)
- Version 2.0.2 (May 2022)
  - Migrated to .NET Standard 2.0 to allow .NET Core projects to consume
  - Added FindPages API to OneNoteProvider, and OpenInOneNote method to OneNotePage (#2, #6)
  - A couple bug fixes and improvements in the OneNote project and the unit tests (#5)
- Version 2.0 (June 2017)
  - Migrated to .NET 4.6
  - Fixed some bugfixes and implemented small improvements
  - Removed factory method of ExcelProvider and removed ExcelVersion enum
  - Moved all classes to subfolders with namespaces Excel, Outlook and OneNote
  - Added references to NuGet packages of OneNote interop and Outlook interop
  - Released 3 new NuGet packages to use ExcelProvider, OutlookProvider and OneNoteProvider standalone
  - Added LinqPad examples
- Version 1.3 (December 2008): .NET 3.5
- Version 1.2 (October 2008): .NET 3.5
- Version 1.1 (December 2007): .NET 3.5
- Version 1.0 (November 2007): .NET 3.5

License
=================================================================

- Released under MIT license