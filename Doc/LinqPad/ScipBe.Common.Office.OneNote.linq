<Query Kind="Program">
  <NuGetReference>ScipBe.Common.Office.OneNote</NuGetReference>
  <Namespace>Microsoft.Office.Interop.OneNote</Namespace>
  <Namespace>ScipBe.Common.Office.OneNote</Namespace>
</Query>

void Main()
{
    Util.RawHtml($"<h2>ScipBe.Common.Office.OneNote - LINQ to OneNote - Query Notebooks, Sections and Pages</h2>").Dump();
	
    // Find pages containing the word "onenote"
    OneNoteProvider.FindPages("onenote")
    .Select(p => new { NotebookName = p.Notebook.Name, SectionName = p.Section.Name, p.Name })
    .Dump("All OneNote Pages containing the word \"onenote\"");
    
	// Show all encrypted OneNote Notebooks
	var queryEncryptedSections = 
	from nb in OneNoteProvider.NotebookItems
	from s in nb.Sections
	where s.Encrypted == true
	select new { NotebookName = nb.Name, SectionName = s.Name };
	queryEncryptedSections.Dump("All encrypted OneNote Notebooks");
	
	// Show all OneNote Notebooks and the number of Sections they have
	var queryNotebooks = 
	(from nb in OneNoteProvider.NotebookItems
	select new { Notebook = nb.Name, SectionCount = nb.Sections.Count() })
	.OrderByDescending(n => n.SectionCount);	
	queryNotebooks.Dump("All OneNote Notebooks and the number of Sections they have");

	// Show all OneNote Pages which have been modified last few days
	var queryPages = 
	from page in OneNoteProvider.PageItems
	where page.LastModified > DateTime.Now.AddMonths(-2)	
	orderby page.LastModified descending
	select new { NotebookName = page.Notebook.Name, SectionName = page.Section.Name, page.Name, page.LastModified };
	queryPages.Dump("All OneNote Pages which have been modified last few days");	
	
	// Show XML content of the OneNote Pages which have been changed the last few days
	foreach (var item in OneNoteProvider.PageItems.Where(p => p.LastModified > DateTime.Now.AddDays(-2)))
	{
        item.GetContent().Dump($"{item.LastModified} {item.Notebook.Name} {item.Section.Name} {item.Name} {item.DateTime}");
	}		
}