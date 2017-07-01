<Query Kind="Program">
  <NuGetReference>ScipBe.Common.Office.Outlook</NuGetReference>
  <Namespace>Microsoft.Office.Interop.Outlook</Namespace>
  <Namespace>ScipBe.Common.Office</Namespace>
  <Namespace>ScipBe.Common.Office.Outlook</Namespace>
</Query>

void Main()
{
    Util.RawHtml($"<h2>ScipBe.Common.Office.Outlook - LINQ to Outlook - Query Appointments</h2>").Dump();
    
	var outlookProvider = new OutlookProvider();
	
	var apointmentItems =  outlookProvider.CalendarItems;
    //var apointmentItems =  outlookProvider.GetItems<AppointmentItem>(@"\\account\Agenda");

    var query = 
    from a in apointmentItems
    where a.Start >= DateTime.Today.AddMonths(-1)
    //where a.IsRecurring
    //where a.Subject != null
    //where a.Subject.Contains("Brown bag session")
    //where a.Categories.Contains("Sport")
    //where a.Start != a.End
    select new
    {
        a.Start,
        a.End,
        a.Duration,
        a.Subject,
        a.IsRecurring,
        a.ReminderSet,
        a.ReminderMinutesBeforeStart,
        a.Organizer,
        a.Location,
        a.Categories
    };

    query.Dump("Outlook Appointments");
}