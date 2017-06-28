using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

namespace ScipBe.Common.Office.Outlook
{
    public interface IOutlookProvider
    {
        /// <summary>
        /// Instance of Outlook Application object.
        /// </summary>
        Application Outlook { get; }

        /// <summary>
        /// Collection of Folder items.
        /// The Outlook class holds a hierarchical structure of folders.
        /// This property will flatten this hierarchy and return a collection with all folders.
        /// </summary>
        IEnumerable<Folder> Folders { get; }

        /// <summary>
        /// Get Outlook items of a given folder.
        /// </summary>
        /// <typeparam name="T">Type (AppointmentItem, ContactItem, JournalItem, MailItem, NoteItem, PostItem, TaskItem).</typeparam>
        /// <param name="folder">Outlook folder.</param>
        IEnumerable<T> GetItems<T>(Folder folder);

        /// <summary>
        /// Get Outlook items of a given folder path.
        /// </summary>
        /// <typeparam name="T">Type (AppointmentItem, ContactItem, JournalItem, MailItem, NoteItem, PostItem, TaskItem).</typeparam>
        /// <param name="folderPath">Outlook folder path.</param>
        /// <returns>Collection of items (AppointmentItems, ContactItems, MailItems, ...).</returns>
        IEnumerable<T> GetItems<T>(string folderPath);

        /// <summary>
        /// Collection of Contact Items of default Contacts folder.
        /// </summary>
        IEnumerable<ContactItem> ContactItems { get; }

        /// <summary>
        /// Collection of Appointment (Calendar) Items of default Calendar.
        /// </summary>
        IEnumerable<AppointmentItem> CalendarItems { get; }

        /// <summary>
        /// Collection of Mail Items of default Inbox folder.
        /// </summary>
        IEnumerable<MailItem> InboxItems { get; }

        /// <summary>
        /// Collection of Mail Items of default SendMail folder.
        /// </summary>
        IEnumerable<MailItem> SentMailItems { get; }

        /// <summary>
        /// Collection of Note Items of default Notes folder.
        /// </summary>
        IEnumerable<NoteItem> NoteItems { get; }

        /// <summary>
        /// Collection of Task Items of default Tasks folder.
        /// </summary>
        IEnumerable<TaskItem> TaskItems { get; }

    }
}