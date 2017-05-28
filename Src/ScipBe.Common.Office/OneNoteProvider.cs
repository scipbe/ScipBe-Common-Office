// ==============================================================================================
// Namespace   : ScipBe.Common.Office
// Class(es)   : OneNoteProvider (LINQ to OneNote)
// Version     : 1.3
// Author      : Stefan Cruysberghs
// Website     : http://www.scip.be
// Date        : October 2007 - November 2008
// Description : Wrapper class for OneNote Xml hierarchy data (notebooks, sections, pages)
// Status      : Open source - MIT License
// ==============================================================================================

using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace ScipBe.Common.Office
{
  /// <summary>
  /// OneNote Provider (LINQ to OneNote) 
  /// </summary>
  /// <remarks>
  /// <list type="bullet">
  /// <item>Author: Stefan Cruysberghs</item>
  /// <item>Website: http://www.scip.be</item>
  /// <item>Article: Querying Outlook and OneNote with LINQ : http://www.scip.be/index.php?Page=ArticlesNET05</item>
  /// <item>Source code: http://www.scip.be/index.php?Page=ComponentsNETOfficeItems</item>
  /// </list>
  /// </remarks>  
  public class OneNoteProvider
  {
    private readonly Microsoft.Office.Interop.OneNote.Application oneNote;

    /// <summary>
    /// Instance of OneNote Application object
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>http://msdn.microsoft.com/en-us/library/ms788684.aspx</item>
    /// <item>http://msdn.microsoft.com/en-us/library/aa286798.aspx</item>
    /// </list>
    /// </remarks>
    public Microsoft.Office.Interop.OneNote.Application OneNote
    {
      get { return oneNote; }
    }

    /// <summary>
    /// Hierarchy of Notebooks with Sections and Pages
    /// </summary>
    /// <example><code>
    /// // All sections which are encrypted
    /// var queryEncryptedSections = from nb in oneNoteProvider.NotebookItems
    ///                              from s in nb.Sections
    ///                              where s.Encrypted == true
    ///                              select new { NotebookName = nb.Name, SectionName = s.Name };
    ///
    /// // All notebooks and the number of sections they have
    /// var queryNotebooks = (from nb in oneNoteProvider.NotebookItems
    ///                       select new { Notebook = nb.Name, SectionCount = nb.Sections.Count() })
    ///                      .OrderByDescending(n =&gt; n.SectionCount);
    /// </code></example>     
    public IEnumerable<IOneNoteExtNotebook> NotebookItems
    {
      get
      {
        XElement oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
        XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

        // Transform XML into object hierarchy
        IEnumerable<OneNoteExtNotebook> oneNoteNotebookItems = from n in oneNoteHierarchy.Elements(one + "Notebook")
                                                               where n.HasAttributes
                                                               select new OneNoteExtNotebook()
                                                               {
                                                                 ID = n.Attribute("ID").Value,
                                                                 Name = n.Attribute("name").Value,
                                                                 NickName = n.Attribute("nickname").Value,
                                                                 Path = n.Attribute("path").Value,
                                                                 Color = ColorTranslator.FromHtml(n.Attribute("color").Value),
                                                                 Sections = n.Elements().Select(s => new OneNoteExtSection()
                                                                 {
                                                                   ID = s.Attribute("ID").Value,
                                                                   Name = s.Attribute("name").Value,
                                                                   Path = s.Attribute("path").Value,
                                                                   Color = ColorTranslator.FromHtml(s.Attribute("color").Value),
                                                                   Encrypted = ((s.Attribute("encrypted") != null) && (s.Attribute("encrypted").Value == "true")),
                                                                   Pages = s.Elements().Select(p => new OneNotePage()
                                                                   {
                                                                     ID = p.Attribute("ID").Value,
                                                                     Name = p.Attribute("name").Value,
                                                                     DateTime = XmlConvert.ToDateTime(p.Attribute("dateTime").ToString(), XmlDateTimeSerializationMode.Utc),
                                                                     LastModified = XmlConvert.ToDateTime(p.Attribute("lastModifiedTime").ToString(), XmlDateTimeSerializationMode.Utc),
                                                                   }).OfType<IOneNotePage>()
                                                                 }).OfType<IOneNoteExtSection>()
                                                               };

        return oneNoteNotebookItems.OfType<IOneNoteExtNotebook>();
      }
    }

    /// <summary>
    /// Collection of Pages
    /// </summary>
    /// <example><code>
    /// // Query all pages which have been modified last month
    /// var queryPages = from page in oneNoteProvider.PageItems
    ///                  where page.LastModified &gt; DateTime.Now.AddMonths(-1)
    ///                  orderby page.LastModified descending
    ///                  select page;
    /// 
    /// // Show XML content of pages which have been changed yesterday
    /// foreach (var item in oneNoteProvider.PageItems.Where(p =&gt; p.LastModified &gt; DateTime.Now.AddDays(-1)))
    /// {
    ///   string pageXMLContent = "";
    ///   Console.WriteLine("{0} {1} {2} {3} {4}", item.LastModified, item.Notebook.Name, item.Section.Name, item.Name, item.DateTime);
    ///   oneNoteProvider.OneNote.GetPageContent(item.ID, out pageXMLContent, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
    ///   Console.WriteLine("{0}", pageXMLContent);
    /// }
    /// </code></example> 
    public IEnumerable<IOneNoteExtPage> PageItems
    {
      get
      {
        XElement oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
        XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

        // Transform XML into object collection
        IEnumerable<OneNoteExtPage> oneNotePageItems = from o in oneNoteHierarchy.Elements(one + "Notebook").Elements().Elements()
                                                       where o.HasAttributes
                                                       select new OneNoteExtPage()
                                                       {
                                                         ID = o.Attribute("ID").Value,
                                                         Name = o.Attribute("name").Value,
                                                         DateTime = XmlConvert.ToDateTime(o.Attribute("dateTime").Value, XmlDateTimeSerializationMode.Utc),
                                                         LastModified = XmlConvert.ToDateTime(o.Attribute("lastModifiedTime").Value, XmlDateTimeSerializationMode.Utc),
                                                         Section = new OneNoteSection()
                                                         {
                                                           ID = o.Parent.Attribute("ID").Value,
                                                           Name = o.Parent.Attribute("name").Value,
                                                           Path = o.Parent.Attribute("path").Value,
                                                           Color = ColorTranslator.FromHtml(o.Parent.Attribute("color").Value),
                                                           Encrypted = ((o.Parent.Attribute("encrypted") != null) && (o.Parent.Attribute("encrypted").Value == "true"))
                                                         },
                                                         Notebook = new OneNoteNotebook()
                                                         {
                                                           ID = o.Parent.Parent.Attribute("ID").Value,
                                                           Name = o.Parent.Parent.Attribute("name").Value,
                                                           NickName = o.Parent.Parent.Attribute("nickname").Value,
                                                           Path = o.Parent.Parent.Attribute("path").Value,
                                                           Color = ColorTranslator.FromHtml(o.Parent.Parent.Attribute("color").Value)
                                                         }
                                                       };

        return oneNotePageItems.OfType<IOneNoteExtPage>();
      }
    }

    private readonly string oneNoteXMLHierarchy = "";

    /// <summary>
    /// Constructor. Create instance of Microsoft.Office.Interop.OneNote.Application and get XML hierarchy.
    /// </summary>
    public OneNoteProvider()
    {
      oneNote = new Microsoft.Office.Interop.OneNote.Application();
      // Get OneNote hierarchy as XML document
      oneNote.GetHierarchy(null, Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, out oneNoteXMLHierarchy);
    }
  }
}
