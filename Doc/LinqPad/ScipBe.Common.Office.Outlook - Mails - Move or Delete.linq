<Query Kind="Program">
  <NuGetReference>ScipBe.Common.Office.Outlook</NuGetReference>
  <Namespace>Microsoft.Office.Interop.Outlook</Namespace>
  <Namespace>ScipBe.Common.Office.Outlook</Namespace>
</Query>

void Main()
{
    Util.RawHtml($"<h2>ScipBe.Common.Office.Outlook - LINQ to Outlook - Move or Delete Mails</h2>").Dump();
    
	var outlookProvider = new OutlookProvider();
	
	var mailItems =  outlookProvider.InboxItems;
    //var mailItems =  outlookProvider.SentMailItems;
    //var mailItems =  outlookProvider.GetItems<AppointmentItem>(@"\\account\Mailbox IN");
   
    var toMailFolder =  outlookProvider.Folders.FirstOrDefault(f => f.FolderPath == @"\\account\Mailbox Archive");

    var oldMails = mailItems.Where(m => m.ReceivedTime >= DateTime.Today.AddYears(-1));

    while (oldMails.Count() > 0)
    {
        var mail = oldMails.Last();
        //mail.Delete();
        //mail.Move(toMailFolder);
    }
}