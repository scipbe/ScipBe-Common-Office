using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using ScipBe.Common.Office.Utils;

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
        /// <summary>
        /// Constructor. Create instance of Microsoft.Office.Interop.OneNote.Application and get XML hierarchy.
        /// </summary>
        public OneNoteProvider()
        {
            this.OneNote = new Microsoft.Office.Interop.OneNote.Application();
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
        public Microsoft.Office.Interop.OneNote.Application OneNote { get; private set; }

        /// <summary>
        /// Hierarchy of Notebooks with Sections and Pages.
        /// </summary>
        public IEnumerable<IOneNoteExtNotebook> NotebookItems
        {
            get
            {
                // Get OneNote hierarchy as XML document
                this.OneNote.GetHierarchy(null, Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, out string oneNoteXMLHierarchy);

                var oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
                var one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

                // Transform XML into object hierarchy
                return from n in oneNoteHierarchy.Elements(one + "Notebook")
                       where n.HasAttributes
                       select this.ParseNotebook(n, one, true);
            }
        }

        /// <summary>
        /// Collection of Pages.
        /// </summary>
        public IEnumerable<IOneNoteExtPage> PageItems
        {
            get
            {
                // Get OneNote hierarchy as XML document
                this.OneNote.GetHierarchy(null, Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, out string oneNoteXMLHierarchy);

                return this.ParsePages(oneNoteXMLHierarchy);
            }
        }

        /// <summary>
        /// Returns a list of pages that match the specified query term.
        /// </summary>
        /// <param name="searchString">The search string. Pass exactly the same string that you would type into the search box in the OneNote UI. You can use bitwise operators, such as AND and OR, which must be all uppercase.</param>
        public IEnumerable<IOneNoteExtPage> FindPages(string searchString)
        {
            this.OneNote.FindPages(null, searchString, out string xml);

            return this.ParsePages(xml);
        }

        private IOneNoteExtNotebook ParseNotebook(XElement element, XNamespace oneNamespace, bool addSections)
        {
            var notebook = new OneNoteExtNotebook()
            {
                ID = element.Attribute("ID").Value,
                Name = element.Attribute("name").Value,
                NickName = element.Attribute("nickname").Value,
                Path = element.Attribute("path").Value,
                Color = element.Attribute("color").Value != "none" ? ColorTranslator.FromHtml(element.Attribute("color").Value) : (Color?)null,
            };

            if (addSections)
            {
                notebook.Sections = element.Elements(oneNamespace + "Section").Select(s => this.ParseSection(s, oneNamespace, true));
            }

            return notebook;
        }

        private IOneNoteExtSection ParseSection(XElement element, XNamespace oneNamespace, bool addPages)
        {
            var section = new OneNoteExtSection()
            {
                ID = element.Attribute("ID").Value,
                Name = element.Attribute("name").Value,
                Path = element.Attribute("path").Value,
                Color = element.Attribute("color").Value != "none" ? ColorTranslator.FromHtml(element.Attribute("color").Value) : (Color?)null,
                Encrypted = (element.Attribute("encrypted") != null) && (element.Attribute("encrypted").Value == "true"),
            };

            if (addPages)
            {
                section.Pages = element.Elements(oneNamespace + "Page").Select(p => this.ParsePage(p, oneNamespace, false));
            }

            return section;
        }

        private IOneNoteExtPage ParsePage(XElement element, XNamespace oneNamespace, bool addParents)
        {
            var page = new OneNoteExtPage(this.OneNote)
            {
                ID = element.Attribute("ID").Value,
                Name = element.Attribute("name").Value,
                Level = element.Attribute("pageLevel").Value.ToInt32(),
                DateTime = element.Attribute("dateTime").Value.ToString().ToDateTime(),
                LastModified = element.Attribute("lastModifiedTime").Value.ToString().ToDateTime(),
            };

            if (addParents)
            {
                page.Section = this.ParseSection(element.Parent, oneNamespace, false);
                page.Notebook = this.ParseNotebook(element.Parent.Parent, oneNamespace, false);
            }

            return page;
        }

        private IEnumerable<IOneNoteExtPage> ParsePages(string xml)
        {
            var doc = XElement.Parse(xml);
            var one = doc.GetNamespaceOfPrefix("one");

            // Transform XML into object collection
            return from p in doc.Elements(one + "Notebook").Elements().Elements()
                   where p.HasAttributes
                   && p.Name.LocalName == "Page"
                   select this.ParsePage(p, one, true);
        }
    }
}
