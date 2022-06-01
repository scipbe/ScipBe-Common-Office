using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
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
    public static class OneNoteProvider
    {
        /// <summary>
        /// Hierarchy of Notebooks with Sections and Pages.
        /// </summary>
        public static IEnumerable<IOneNoteExtNotebook> NotebookItems
        {
            get
            {
                return CallOneNoteSafely(oneNote =>
                {
                    // Get OneNote hierarchy as XML document
                    oneNote.GetHierarchy(null, HierarchyScope.hsPages, out string oneNoteXMLHierarchy);
                    var oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
                    var one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

                    // Transform XML into object hierarchy
                    return from n in oneNoteHierarchy.Elements(one + "Notebook")
                           where n.HasAttributes
                           select ParseNotebook(n, one, true);
                });
            }
        }

        /// <summary>
        /// Collection of Pages.
        /// </summary>
        public static IEnumerable<IOneNoteExtPage> PageItems
        {
            get
            {
                return CallOneNoteSafely(oneNote =>
                {
                    // Get OneNote hierarchy as XML document
                    oneNote.GetHierarchy(null, HierarchyScope.hsPages, out string oneNoteXMLHierarchy);
                    return ParsePages(oneNoteXMLHierarchy);
                });
            }
        }

        /// <summary>
        /// Returns a list of pages that match the specified query term.
        /// </summary>
        /// <param name="searchString">The search string. Pass exactly the same string that you would type into the search box in the OneNote UI. You can use bitwise operators, such as AND and OR, which must be all uppercase.</param>
        public static IEnumerable<IOneNoteExtPage> FindPages(string searchString)
        {
            return CallOneNoteSafely(oneNote =>
            {
                oneNote.FindPages(null, searchString, out string xml);
                return ParsePages(xml);
            });
        }

        private static IOneNoteExtNotebook ParseNotebook(XElement element, XNamespace oneNamespace, bool addSections)
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
                notebook.Sections = element.Descendants(oneNamespace + "Section").Select(s => ParseSection(s, oneNamespace, true));
            }

            return notebook;
        }

        private static IOneNoteExtSection ParseSection(XElement element, XNamespace oneNamespace, bool addPages)
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
                section.Pages = element.Elements(oneNamespace + "Page").Select(p => ParsePage(p, oneNamespace, false));
            }

            return section;
        }

        private static IOneNoteExtPage ParsePage(XElement element, XNamespace oneNamespace, bool addParents)
        {
            var page = new OneNoteExtPage()
            {
                ID = element.Attribute("ID").Value,
                Name = element.Attribute("name").Value,
                Level = element.Attribute("pageLevel").Value.ToInt32(),
                DateTime = element.Attribute("dateTime").Value.ToString().ToDateTime(),
                LastModified = element.Attribute("lastModifiedTime").Value.ToString().ToDateTime(),
            };

            if (addParents)
            {
                page.Section = ParseSection(element.Parent, oneNamespace, false);

                var notebookElement = element.Parent.Parent;
                if (notebookElement.Name.LocalName == "SectionGroup")
                {
                    notebookElement = notebookElement.Parent;
                }

                page.Notebook = ParseNotebook(notebookElement, oneNamespace, false);
            }

            return page;
        }

        public static IEnumerable<IOneNoteExtPage> ParsePages(string xml)
        {
            var doc = XElement.Parse(xml);
            var one = doc.GetNamespaceOfPrefix("one");

            var sections = from section in doc.Elements(one + "Notebook").Descendants(one + "Section")
                           select section;

            // Transform XML into object collection
            return from p in sections.Elements()
                   where p.HasAttributes
                   && p.Name.LocalName == "Page"
                   select ParsePage(p, one, true);
        }

        internal static T CallOneNoteSafely<T>(Func<Application, T> action)
        {
            Application oneNote = null;
            try
            {
                oneNote = Util.TryCatchAndRetry<Application, COMException>(
                    () => new Application(),
                    TimeSpan.FromMilliseconds(100),
                    3,
                    ex => Trace.TraceError(ex.Message));
                return action(oneNote);
            }
            finally
            {
                if (oneNote != null)
                {
                    Marshal.ReleaseComObject(oneNote);
                }
            }
        }
    }
}
