using System.Collections.Generic;
using Microsoft.Office.Interop.OneNote;

namespace ScipBe.Common.Office.OneNote
{
    public interface IOneNoteProvider
    {
        /// <summary>
        /// Instance of OneNote Application object.
        /// </summary>
        Application OneNote { get; }

        /// <summary>
        /// Hierarchy of Notebooks with Sections and Pages.
        /// </summary>
        IEnumerable<IOneNoteExtNotebook> NotebookItems { get; }

        /// <summary>
        /// Collection of Pages.
        /// </summary>
        IEnumerable<IOneNoteExtPage> PageItems { get; }

        /// <summary>
        /// Returns a list of pages that match the specified query term.
        /// </summary>
        /// <param name="searchString">The search string. Pass exactly the same string that you would type into the search box in the OneNote UI.
        /// You can use bitwise operators, such as AND and OR, which must be all uppercase.</param>
        IEnumerable<IOneNoteExtPage> FindPages(string searchString);
    }
}