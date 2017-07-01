<Query Kind="Program">
  <NuGetReference>ScipBe.Common.Office.Outlook</NuGetReference>
  <Namespace>Microsoft.Office.Interop.Outlook</Namespace>
  <Namespace>ScipBe.Common.Office.Outlook</Namespace>
</Query>

void Main()
{
    Util.RawHtml($"<h2>ScipBe.Common.Office.Outlook - LINQ to Outlook - Remove large Attachments from Mails</h2>").Dump();
    
	var outlookProvider = new OutlookProvider();

    var mails = outlookProvider.InboxItems;
    //var mails = outlookProvider.SentMailItems;
    //var mails = outlookProvider.GetItems<MailItem>(@"\\account\Mailbox OUT");

    var maxSizeInKb = 1024 * 500;

    var mailsWithAttachments = 
    from m in mails
    let a = m.Attachments.OfType<Attachment>()
    where
    m.Attachments.Count > 0
    && m.Attachments.OfType<Attachment>().Any(a => a.Size > maxSizeInKb)
    select m;

    ((mailsWithAttachments.Sum(m => m.Attachments.OfType<Attachment>().Sum(a => a.Size)) / 1024 / 1024) + "MB").Dump("Total size of attachments which will be deleted");

    foreach (var mail in mailsWithAttachments)
    {
        foreach (var attachment in mail.Attachments.OfType<Attachment>().Where(a => a.Size > maxSizeInKb))
        {
            $"{mail.CreationTime} - {mail.Subject} - {attachment.DisplayName} - {attachment.Size / 1024}KB".Dump();
            mail.Attachments.Remove(attachment.Index);
            //mail.Save();
        }
    }
}