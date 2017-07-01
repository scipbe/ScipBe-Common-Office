<Query Kind="Program">
  <NuGetReference>ScipBe.Common.Office.Outlook</NuGetReference>
  <Namespace>Microsoft.Office.Interop.Outlook</Namespace>
  <Namespace>ScipBe.Common.Office.Outlook</Namespace>
</Query>

void Main()
{
    Util.RawHtml($"<h2>ScipBe.Common.Office.Outlook - LINQ to Outlook - Statistics about mails</h2>").Dump();
    
	var outlookProvider = new OutlookProvider();

    var mails = outlookProvider.InboxItems;
    //var mails = outlookProvider.SentMailItems;
    //var mails = outlookProvider.GetItems<MailItem>(@"\\account\Mailbox OUT");

    var query = 
    from m in mails
    orderby m.SentOn descending
    group m by $"{m.SentOn.Date.Year}-{m.SentOn.Date.Month}" into g
    select new
    {
        SentOn = g.Key,
        Count = g.Count()
    };
    
    query.Dump("Mail statistics");
}