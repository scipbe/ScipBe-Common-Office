using ScipBe.Common.Office.Utils;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;

namespace ScipBe.Common.Office.OneNote
{
    /// <summary>
    /// OneNote Provider (LINQ to OneNote).
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>Author: Stefan Cruysberghs</item>
    /// <item>Website: http://www.scip.be</item>
    /// <item>Article: Querying Outlook and OneNote with LINQ : http://www.scip.be/index.php?Page=ArticlesNET05</item>
    /// </list>
    /// </remarks>
    public class OneNoteProvider : IOneNoteProvider
    {
        private readonly Microsoft.Office.Interop.OneNote.Application oneNote;

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

        /// <summary>
        /// Instance of OneNote Application object.
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
        /// Hierarchy of Notebooks with Sections and Pages.
        /// </summary>
        public IEnumerable<IOneNoteExtNotebook> NotebookItems
        {
            get
            {
                var oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
                var one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

                // Transform XML into object hierarchy
                var oneNoteNotebookItems =
                    from n in oneNoteHierarchy.Elements(one + "Notebook")
                    where n.HasAttributes
                    select new OneNoteExtNotebook()
                    {
                        ID = n.Attribute("ID").Value,
                        Name = n.Attribute("name").Value,
                        NickName = n.Attribute("nickname").Value,
                        Path = n.Attribute("path").Value,
                        Color = ColorTranslator.FromHtml(n.Attribute("color").Value),
                        Sections = n.Elements(one + "Section").Select(s => new OneNoteExtSection()
                        {
                            ID = s.Attribute("ID").Value,
                            Name = s.Attribute("name").Value,
                            Path = s.Attribute("path").Value,
                            Color = ColorTranslator.FromHtml(s.Attribute("color").Value),
                            Encrypted = ((s.Attribute("encrypted") != null) && (s.Attribute("encrypted").Value == "true")),
                            Pages = s.Elements(one + "Page").Select(p => new OneNotePage()
                            {
                                ID = p.Attribute("ID").Value,
                                Name = p.Attribute("name").Value,
                                Level = p.Attribute("pageLevel").Value.ToInt32(),
                                DateTime = p.Attribute("dateTime").Value.ToString().ToDateTime(),
                                LastModified = p.Attribute("lastModifiedTime").Value.ToString().ToDateTime(),
                            }).OfType<IOneNotePage>()
                        }).OfType<IOneNoteExtSection>()
                    };

                return oneNoteNotebookItems.OfType<IOneNoteExtNotebook>();
            }
        }

        /// <summary>
        /// Collection of Pages.
        /// </summary>
        public IEnumerable<IOneNoteExtPage> PageItems
        {
            get
            {
                var oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
                var one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

                // Transform XML into object collection
                var oneNotePageItems =
                    from p in oneNoteHierarchy.Elements(one + "Notebook").Elements().Elements()
                    where p.HasAttributes
                    && p.Name.LocalName == "Page"
                    select new OneNoteExtPage()
                    {
                        ID = p.Attribute("ID").Value,
                        Name = p.Attribute("name").Value,
                        Level = p.Attribute("pageLevel").Value.ToInt32(),
                        DateTime = p.Attribute("dateTime").Value.ToString().ToDateTime(),
                        LastModified = p.Attribute("lastModifiedTime").Value.ToString().ToDateTime(),
                        Section = new OneNoteSection()
                        {
                            ID = p.Parent.Attribute("ID").Value,
                            Name = p.Parent.Attribute("name").Value,
                            Path = p.Parent.Attribute("path").Value,
                            Color = ColorTranslator.FromHtml(p.Parent.Attribute("color").Value),
                            Encrypted = ((p.Parent.Attribute("encrypted") != null) && (p.Parent.Attribute("encrypted").Value == "true"))
                        },
                        Notebook = new OneNoteNotebook()
                        {
                            ID = p.Parent.Parent.Attribute("ID").Value,
                            Name = p.Parent.Parent.Attribute("name").Value,
                            NickName = p.Parent.Parent.Attribute("nickname").Value,
                            Path = p.Parent.Parent.Attribute("path").Value,
                            Color = ColorTranslator.FromHtml(p.Parent.Parent.Attribute("color").Value)
                        }
                    };

                return oneNotePageItems.OfType<IOneNoteExtPage>();
            }
        }
    }
}
