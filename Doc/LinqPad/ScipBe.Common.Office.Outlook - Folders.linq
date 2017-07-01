<Query Kind="Program">
  <NuGetReference>ScipBe.Common.Office.Outlook</NuGetReference>
  <Namespace>Microsoft.Office.Interop.Outlook</Namespace>
  <Namespace>ScipBe.Common.Office.Outlook</Namespace>
</Query>

void Main()
{
    Util.RawHtml($"<h2>ScipBe.Common.Office.Outlook - LINQ to Outlook - Query folders</h2>").Dump();
    
	var outlookProvider = new OutlookProvider();

	outlookProvider.Folders.Select(f => 
	new
	{
		f.FolderPath,
		f.Name,
		f.DefaultItemType,
		f.UnReadItemCount,
		SubfoldersCount = f.Folders.Count,
		ItemsCount = f.Items.Count,
	}).Dump("All Outlook folders");
}