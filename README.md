![Logo](Doc/Images/ScipBe.png) 
# ScipBe-Common-Office
### Linq to Excel, Outlook and OneNote

The ScipBe.Common.Office namespace contains 3 classes : ExcelProvider (LINQ to Excel), OutlookProvider (LINQ to Outlook) and OneNoteProvider (LINQ to OneNote). 
- The ExcelProvider loads an Excel worksheet and provides column definition and row collections. All collections are IEnumerable so you can query them with LINQ. 
- The OutlookProvider is a wrapper class which provides IEnumerable collections to data of the COM interface of Outlook (appointments, contacts, mails, tasks, ...). 
- The OneNoteProvider provides collections of notebooks, sections and pages by parsing the XML hierarchy tree of OneNote. 

Links
=================================================================

- [Homepage](http://www.scip.be)
- [Documentation and examples](http://www.scip.be/index.php?Page=ComponentsNETOfficeItems)
- [Author Stefan Cruysberghs](http://www.scip.be/index.php?Page=AboutMe)

History
=================================================================

- Version 1.3 (December 2008): .NET 3.5
- Version 1.2 (October 2008): .NET 3.5
- Version 1.1 (December 2007): .NET 3.5
- Version 1.0 (November 2007): .NET 3.5

License
=================================================================

- Released under MIT license
