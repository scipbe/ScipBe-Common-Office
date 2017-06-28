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
    }
}