using Microsoft.Office.Interop.Outlook;
using System;

namespace ScipBe.Common.Office.Outlook
{
    public static class OlItemTypeExtensions
    {
        public static Type GetItemType(this OlItemType olItemType)
        {
            switch (olItemType)
            {
                case OlItemType.olMailItem:
                    return typeof(MailItem);
                case OlItemType.olAppointmentItem:
                    return typeof(AppointmentItem);
                case OlItemType.olContactItem:
                    return typeof(ContactItem);
                case OlItemType.olTaskItem:
                    return typeof(TaskItem);
                case OlItemType.olJournalItem:
                    return typeof(JournalItem);
                case OlItemType.olNoteItem:
                    return typeof(NoteItem);
                case OlItemType.olPostItem:
                    return typeof(PostItem);
                case OlItemType.olDistributionListItem:
                    return typeof(DistListItem);
                case OlItemType.olMobileItemSMS:
                    return typeof(MobileItem);
                case OlItemType.olMobileItemMMS:
                    return typeof(MobileItem);
                default:
                    return null;
            }
        }
    }
}
