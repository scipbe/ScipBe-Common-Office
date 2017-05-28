// ==============================================================================================
// Namespace   : ScipBe.Common.Office
// Class(es)   : OutlookProvider (LINQ to Outlook)
// Version     : 1.4
// Author      : Stefan Cruysberghs
// Website     : http://www.scip.be
// Date        : October 2007 - January 2009
// Description : Wrapper class for Outlook data (mails, contacts, appointments, notes, tasks)
// Status      : Open source - MIT License
// ==============================================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Outlook;

namespace ScipBe.Common.Office
{
  /// <summary>
  /// Outlook Provider (LINQ to Outlook) 
  /// </summary>
  /// <remarks>
  /// <list type="bullet">
  /// <item>Author: Stefan Cruysberghs</item>
  /// <item>Website: http://www.scip.be</item>
  /// <item>Article: Querying Outlook and OneNote with LINQ : http://www.scip.be/index.php?Page=ArticlesNET05</item>
  /// <item>Article: Execute queries on Office data with LINQPad : http://www.scip.be/index.php?Page=ArticlesNET06</item>
  /// <item>Article: Display Outlook contact pictures in WPF application : http://www.scip.be/index.php?Page=ArticlesNET07</item>
  /// <item>Article: Cleaning up Outlook mailboxes : http://www.scip.be/index.php?Page=ArticlesNET27</item>
  /// <item>Source code: http://www.scip.be/index.php?Page=ComponentsNETOfficeItems</item>
  /// </list>
  /// </remarks>
  public class OutlookProvider
  {
    private readonly Application outlook;

    /// <summary>
    /// Constructor. Create instance of Office.Interop.Outlook.Application
    /// </summary>
    public OutlookProvider()
    {
      outlook = new Application();
    }

    /// <summary>
    /// Instance of Outlook Application object
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
        return GetAllFolders(outlook.ActiveExplorer().Session.Folders.OfType<Folder>()); 
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
    /// Get Outlook items of a given folder path
    /// </summary>
    /// <typeparam name="T">Type (AppointmentItem, ContactItem, JournalItem, MailItem, NoteItem, PostItem, TaskItem)</typeparam>
    /// <param name="path">Folder path</param>
    /// <returns>Collection of items (appointments, contactitems, mailitems, ...)</returns>
    /// <example><code>
    /// var appointments = outlookProvider.GetItems&lt;AppointmentItem&gt;(@"\\Personal folders\Agenda");
    /// var mails = outlookProvider.GetItems&lt;MailItem&gt;(@"\\Archive\Sent mails\2008");
    /// </code></example>
    public IEnumerable<T> GetItems<T>(string path)
    {
      if (path.Substring(0,2) != @"\\")
      {
          path = @"\\" + path;
      }

      var folder = Folders.FirstOrDefault(f => f.FolderPath == path);

      if (folder == null)
      {
        throw new ArgumentNullException("path", "Path to Outlook folder does not exists");
      }

      if ((folder.DefaultItemType == OlItemType.olAppointmentItem) && (typeof(T) == typeof(AppointmentItem)))
        return folder.Items.OfType<T>();
      if ((folder.DefaultItemType == OlItemType.olContactItem) && (typeof(T) == typeof(ContactItem)))
        return folder.Items.OfType<T>();
      if ((folder.DefaultItemType == OlItemType.olJournalItem) && (typeof(T) == typeof(JournalItem)))
        return folder.Items.OfType<T>();
      if ((folder.DefaultItemType == OlItemType.olMailItem) && (typeof(T) == typeof(MailItem)))
        return folder.Items.OfType<T>();
      if ((folder.DefaultItemType == OlItemType.olNoteItem) && (typeof(T) == typeof(NoteItem)))
        return folder.Items.OfType<T>();
      if ((folder.DefaultItemType == OlItemType.olPostItem) && (typeof(T) == typeof(PostItem)))
        return folder.Items.OfType<T>();
      if ((folder.DefaultItemType == OlItemType.olTaskItem) && (typeof(T) == typeof(TaskItem)))
        return folder.Items.OfType<T>();

      return null;
    }

    /// <summary>
    /// Collection of Contact Items
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._contactitem_properties.aspx</item>
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._contactitem_methods.aspx</item>
    /// </list>
    /// </remarks>    
    /// <example><code>
    /// // Query all contacts which do not have an email address
    /// var queryContacts = from contact in outlookProvider.ContactItems
    ///                     where contact.Email1Address == null
    ///                     select contact;
    /// </code></example>
    public IEnumerable<ContactItem> ContactItems
    {
      get
      {
        //Microsoft.Office.Interop.Outlook.MAPIFolder folder =
        //  outlook.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

        Microsoft.Office.Interop.Outlook.MAPIFolder folder =
          outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderContacts);

        return folder.Items.OfType<ContactItem>();
      }
    }

    /// <summary>
    /// Collection of Appointment (Calendar) Items 
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._appointmentitem_properties.aspx</item>
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._appointmentitem_methods.aspx</item>
    /// </list>
    /// </remarks>    
    /// <example><code>
    /// // Query all sport activities for next month
    /// var queryAppointments = from app in outlookProvider.CalendarItems
    ///                         where app.Start &gt; DateTime.Now
    ///                           &amp;&amp; app.Start &lt; DateTime.Now.AddMonths(1)
    ///                           &amp;&amp; app.Categories != null
    ///                           &amp;&amp; app.Categories.Contains("Sport")
    ///                         select app;
    /// </code></example> 
    public IEnumerable<AppointmentItem> CalendarItems
    {
      get
      {
        //Microsoft.Office.Interop.Outlook.MAPIFolder folder =
        //  outlook.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

        Microsoft.Office.Interop.Outlook.MAPIFolder folder =
            outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

        return folder.Items.OfType<AppointmentItem>();
      }
    }

    /// <summary>
    /// Collection of Inbox (Mail) Items 
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_properties.aspx</item>
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_methods.aspx</item>
    /// </list>
    /// </remarks> 
    public IEnumerable<MailItem> InboxItems
    {
      get
      {
        //Microsoft.Office.Interop.Outlook.MAPIFolder folder =
        //  outlook.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

        Microsoft.Office.Interop.Outlook.MAPIFolder folder =
            outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);

        return folder.Items.OfType<MailItem>();
      }
    }

    /// <summary>
    /// Collection of Sent Mail Items 
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_properties.aspx</item>
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._mailitem_methods.aspx</item>
    /// </list>
    /// </remarks> 
    /// <example><code>
    /// // Query sent mail items. How many mails did I sent each day of last week?
    /// var queryMails = from mail in outlookProvider.SentMailItems
    ///                  where mail.SentOn &gt; DateTime.Now.AddDays(-7)
    ///                  orderby mail.SentOn descending
    ///                  group mail by mail.SentOn.Date into g
    ///                  select new { SentOn = g.First().SentOn.Date, Count = g.Count() };
    ///
    /// // Set flag complete of all sent mails of last week which have at least one attachment
    /// var queryAttMails = from mail in outlookProvider.SentMailItems
    ///                     where mail.SentOn &gt; DateTime.Now.AddDays(-7)
    ///                       &amp;&amp; mail.Attachments.Count &gt; 0
    ///                     select mail;
    /// 
    /// foreach (var item in queryAttMails)
    /// {
    ///   item.FlagStatus = Microsoft.Office.Interop.Outlook.OlFlagStatus.olFlagComplete;
    ///   item.Save();
    /// }
    /// </code></example>
    public IEnumerable<MailItem> SentMailItems
    {
      get
      {
        //Microsoft.Office.Interop.Outlook.MAPIFolder folder =
        //  outlook.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);

        Microsoft.Office.Interop.Outlook.MAPIFolder folder =
            outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderSentMail);

        return folder.Items.OfType<MailItem>();
      }
    }

    /// <summary>
    /// Collection of Note Items 
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._noteitem_properties.aspx</item>
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._noteitem_methods.aspx</item>
    /// </list>
    /// </remarks> 
    public IEnumerable<NoteItem> NoteItems
    {
      get
      {
        //Microsoft.Office.Interop.Outlook.MAPIFolder folder =
        //  outlook.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderNotes);

        Microsoft.Office.Interop.Outlook.MAPIFolder folder =
            outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderNotes);

        return folder.Items.OfType<NoteItem>();
      }
    }

    /// <summary>
    /// Collection of Task Items
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._taskitem_properties.aspx</item>
    /// <item>http://msdn2.microsoft.com/en-us/library/microsoft.office.interop.outlook._taskitem_methods.aspx</item>
    /// </list>
    /// </remarks> 
    /// <example><code>
    /// // Query all tasks which are in progress and have a percentage complete lower then 50
    /// var queryTasks = from task in outlookProvider.TaskItems
    ///                  where task.Status == Microsoft.Office.Interop.Outlook.OlTaskStatus.olTaskInProgress
    ///                    &amp;&amp; task.PercentComplete &lt; 50
    ///                  select task;
    /// 
    /// // Query all contacts which do not have a postalcode
    /// var queryContactsWithoutPostalCode = from contact in outlookProvider.ContactItems
    ///                                      where contact.HomeAddressPostalCode == null
    ///                                        &amp;&amp; contact.HomeAddressCity != null
    ///                                      select contact;
    /// // Concat all names in one string
    /// string contacts = "";
    /// foreach (var item in queryContactsWithoutPostalCode)
    ///   contacts += item.FirstName + " " + item.LastName + "(" + item.HomeAddressCity + "), ";
    /// // Create a new task containing all contact names which do not have a postalcode
    /// Microsoft.Office.Interop.Outlook.TaskItem ti = 
    ///   (Microsoft.Office.Interop.Outlook.TaskItem)outlookProvider.Outlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olTaskItem);
    /// ti.Subject = "Find missing postalcode for contacts";
    /// ti.Status = Microsoft.Office.Interop.Outlook.OlTaskStatus.olTaskNotStarted;
    /// ti.StartDate = DateTime.Now;
    /// ti.Body = contacts;
    /// ti.Save();
    /// </code></example>
    public IEnumerable<TaskItem> TaskItems
    {
      get
      {
        //Microsoft.Office.Interop.Outlook.MAPIFolder folder =
        //  outlook.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderTasks);

        Microsoft.Office.Interop.Outlook.MAPIFolder folder =
            outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderTasks);

        return folder.Items.OfType<TaskItem>();
      }
    }

    /// <summary>
    /// Save contact picture as JPG in Windows temporary folder and return path to this file
    /// </summary>
    /// <param name="contact">Outlook contact item</param>
    /// <returns>Path to JPG file with Outlook contact picture</returns>
    public static string GetContactPicturePath(ContactItem contact)
    {
      return GetContactPicturePath(contact, Path.GetTempPath());
    }

    /// <summary>
    /// Save contact picture as JPG in given folder and return path to this file
    /// </summary>
    /// <param name="contact">Outlook contact item</param>
    /// <param name="path">Path to folder with temporary Outlook contact pictures</param>
    /// <returns>Path to JPG file with Outlook contact picture</returns>
    /// <example><code>
    /// // WPF ValueConverter which returns bitmap for given contact item
    /// public object Convert(object value, Type targetType, object parameter,
    ///   System.Globalization.CultureInfo culture)
    /// {
    ///   // Return null in design time to avoid WPF Designer load failure
    ///   if ((bool)(DesignerProperties.IsInDesignModeProperty.GetMetadata(typeof(DependencyObject)).DefaultValue)) 
    ///   {
    ///     return null;
    ///   }
    ///   else
    ///   {
    ///     BitmapImage bitmap = null;
    ///     Microsoft.Office.Interop.Outlook.ContactItem contact = (Microsoft.Office.Interop.Outlook.ContactItem)value;
    ///     string picturePath = OutlookProvider.GetContactPicturePath(contact);
    ///     if ((picturePath != "") &amp;&amp; (System.IO.File.Exists(picturePath)))
    ///       bitmap = new BitmapImage(new Uri(picturePath, UriKind.Absolute));
    ///     return bitmap;
    ///   }
    /// }
    /// </code></example>
    public static string GetContactPicturePath(ContactItem contact, string path)
    {
      string picturePath = "";

      if (contact.HasPicture)
      {
        foreach (Attachment att in contact.Attachments)
        {
          if (att.DisplayName == "ContactPicture.jpg")
          {
            try
            {
              picturePath = Path.GetDirectoryName(path) + "\\Contact_" + contact.EntryID + ".jpg";
              if (!File.Exists(picturePath))
              {
                  att.SaveAsFile(picturePath);
              }
            }
            catch (IOException ex)
            {
              picturePath = "";
              Debug.Write(ex.Message);
            }
          }
        }
      }

      return picturePath;
    }

    /// <summary>
    /// Cleanup contact pictures in Windows temporary folder
    /// </summary>
    public static void CleanupContactPictures()
    {
      CleanupContactPictures(Path.GetTempPath());
    }

    /// <summary>
    /// Cleanup contact pictures in given folder
    /// </summary>
    /// <param name="path">Path to folder with temporary Outlook contact pictures</param>
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
