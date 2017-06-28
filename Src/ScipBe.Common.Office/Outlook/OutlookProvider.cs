using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Outlook;

namespace ScipBe.Common.Office.Outlook
{
    /// <summary>
    /// Outlook Provider (LINQ to Outlook).
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>Author: Stefan Cruysberghs</item>
    /// <item>Website: http://www.scip.be</item>
    /// <item>Article: Querying Outlook and OneNote with LINQ : http://www.scip.be/index.php?Page=ArticlesNET05</item>
    /// <item>Article: Execute queries on Office data with LINQPad : http://www.scip.be/index.php?Page=ArticlesNET06</item>
    /// <item>Article: Display Outlook contact pictures in WPF application : http://www.scip.be/index.php?Page=ArticlesNET07</item>
    /// <item>Article: Cleaning up Outlook mailboxes : http://www.scip.be/index.php?Page=ArticlesNET27</item>
    /// </list>
    /// </remarks>
    public class OutlookProvider : IOutlookProvider
    {
        private readonly Application outlook;

        /// <summary>
        /// Constructor. Create instance of Office.Interop.Outlook.Application.
        /// </summary>
        public OutlookProvider()
        {
            outlook = new Application();
        }

        /// <summary>
        /// Instance of Outlook Application object.
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._application_properties.aspx</item>
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._application_methods.aspx</item>
        /// </list>
        /// </remarks>
        public Application Outlook
        {
            get { return outlook; }
        }

        /// <summary>
        /// Collection of Folder items.
        /// The Outlook class holds a hierarchical structure of folders.
        /// This property will flatten this hierarchy and return a collection with all folders.
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook._folders_properties.aspx</item>
        /// <item>http://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook._folders_methods.aspx</item>
        /// </list>
        /// </remarks>
        public IEnumerable<Folder> Folders
        {
            get
            {
                return GetAllFolders(outlook.GetNamespace("MAPI").Folders.OfType<Folder>());                
            }
        }

        private static IEnumerable<Folder> GetAllFolders(IEnumerable<Folder> folders)
        {
            foreach (var folder in folders)
            {
                foreach (var subfolder in GetAllFolders(folder.Folders.OfType<Folder>()))
                {
                    yield return subfolder;
                }
                yield return folder;
            }
        }

        /// <summary>
        /// Get Outlook items of a given folder.
        /// </summary>
        /// <typeparam name="T">Type (AppointmentItem, ContactItem, JournalItem, MailItem, NoteItem, PostItem, TaskItem).</typeparam>
        /// <param name="folder">Outlook folder.</param>
        /// <returns>Collection of items (AppointmentItems, ContactItems, MailItems, ...).</returns>
        public IEnumerable<T> GetItems<T>(Folder folder)
        {
            if (folder == null)
            {
                throw new ArgumentNullException(nameof(folder), $"Folder is required");
            }

            var type = folder.DefaultItemType.GetItemType();
            if (type.FullName != typeof(T).FullName)
            {
                throw new ArgumentException($"Folder {folder.FolderPath} does not contain {typeof(T).Name} items");
            }
            return folder.Items.OfType<T>();
        }

        /// <summary>
        /// Get Outlook items of default folder of given type.
        /// </summary>
        /// <typeparam name="T">Type (AppointmentItem, ContactItem, JournalItem, MailItem, NoteItem, PostItem, TaskItem).</typeparam>
        /// <param name="defaultFolderType">Default folder of type.</param>
        /// <returns>Collection of items (AppointmentItems, ContactItems, MailItems, ...).</returns>
        private IEnumerable<T> GetItems<T>(OlDefaultFolders defaultFolderType)
        {
            var folder = outlook.GetNamespace("MAPI").GetDefaultFolder(defaultFolderType);
            return GetItems<T>(folder.FolderPath);
        }

        /// <summary>
        /// Get Outlook items of a given folder path.
        /// </summary>
        /// <typeparam name="T">Type (AppointmentItem, ContactItem, JournalItem, MailItem, NoteItem, PostItem, TaskItem).</typeparam>
        /// <param name="folderPath">Outlook folder path.</param>
        /// <returns>Collection of items (AppointmentItems, ContactItems, MailItems, ...).</returns>
        public IEnumerable<T> GetItems<T>(string folderPath)
        {
            if (folderPath.Substring(0, 2) != @"\\")
            {
                folderPath = @"\\" + folderPath;
            }

            var folder = Folders.FirstOrDefault(f => f.FolderPath == folderPath);

            return GetItems<T>(folder);
        }

        /// <summary>
        /// Collection of Contact Items of default Contacts folder.
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._contactitem_properties.aspx</item>
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._contactitem_methods.aspx</item>
        /// </list>
        /// </remarks>    
        public IEnumerable<ContactItem> ContactItems => GetItems<ContactItem>(OlDefaultFolders.olFolderContacts);

        /// <summary>
        /// Collection of Appointment (Calendar) Items of default Calendar.
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._appointmentitem_properties.aspx</item>
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._appointmentitem_methods.aspx</item>
        /// </list>
        /// </remarks>    
        public IEnumerable<AppointmentItem> CalendarItems => GetItems<AppointmentItem>(OlDefaultFolders.olFolderCalendar);

        /// <summary>
        /// Collection of Mail Items of default Inbox folder.
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_properties.aspx</item>
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_methods.aspx</item>
        /// </list>
        /// </remarks> 
        public IEnumerable<MailItem> InboxItems => outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox).Items.OfType<MailItem>();

        /// <summary>
        /// Collection of Mail Items of default SendMail folder.
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_properties.aspx</item>
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_methods.aspx</item>
        /// </list>
        /// </remarks> 
        public IEnumerable<MailItem> SentMailItems => GetItems<MailItem>(OlDefaultFolders.olFolderSentMail);

        /// <summary>
        /// Collection of Note Items of default Notes folder.
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._noteitem_properties.aspx</item>
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._noteitem_methods.aspx</item>
        /// </list>
        /// </remarks> 
        public IEnumerable<NoteItem> NoteItems => GetItems<NoteItem>(OlDefaultFolders.olFolderNotes);

        /// <summary>
        /// Collection of Task Items of default Tasks folder.
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._taskitem_properties.aspx</item>
        /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._taskitem_methods.aspx</item>
        /// </list>
        /// </remarks> 
        public IEnumerable<TaskItem> TaskItems => GetItems<TaskItem>(OlDefaultFolders.olFolderTasks);

        /// <summary>
        /// Cleanup contact pictures in Windows temporary folder.
        /// </summary>
        public static void CleanupContactPictures()
        {
            CleanupContactPictures(Path.GetTempPath());
        }

        /// <summary>
        /// Cleanup contact pictures in given folder.
        /// </summary>
        /// <param name="path">Path to folder with temporary Outlook contact pictures.</param>
        public static void CleanupContactPictures(string path)
        {
            foreach (string picturePath in Directory.GetFiles(path, "Contact_*.jpg"))
            {
                try
                {
                    File.Delete(picturePath);
                }
                catch (IOException ex)
                {
                    Debug.Write(ex.Message);
                }
            }
        }
    }
}
