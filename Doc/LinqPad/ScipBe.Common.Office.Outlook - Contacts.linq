<Query Kind="Program">
  <NuGetReference>ScipBe.Common.Office.Outlook</NuGetReference>
  <Namespace>Microsoft.Office.Interop.Outlook</Namespace>
  <Namespace>ScipBe.Common.Office</Namespace>
  <Namespace>ScipBe.Common.Office.Outlook</Namespace>
</Query>

void Main()
{
    Util.RawHtml($"<h2>ScipBe.Common.Office.Outlook - LINQ to Outlook - Query Contacts</h2>").Dump();
    
	var outlookProvider = new OutlookProvider();
	
	var contactItems =  outlookProvider.ContactItems;
	//var contactItems =  outlook.GetItems<ContactItem>(@"\\account\Contactpersons");

	var query = 
	from c in contactItems
	//where c.Email1Address == null
	//where c.Attachments.OfType<Attachment>().Count() > 0
	//where c.LastName == ""
	//where (c.Body != null) && (c.Body.Contains("Son of"))
	//where c.Categories != null && c.Categories.Contains("Computerclub")
	select new
	{
		c.FirstName,
		c.LastName,
		c.HomeAddress,
		c.HomeAddressStreet,
		c.HomeAddressPostalCode,
		c.HomeAddressCity,
		c.HomeAddressCountry,
		c.BusinessAddress,
		c.BusinessAddressStreet,
		c.BusinessAddressPostalCode,
		c.BusinessAddressCity,
		c.BusinessAddressCountry,
		c.Email1Address,
		c.Email2Address,
		c.WebPage,
		c.HasPicture,
		c.HomeTelephoneNumber,
		c.MobileTelephoneNumber,
		c.BusinessTelephoneNumber,
		c.Categories
	};

	query.Dump("Outlook Contacts");
}