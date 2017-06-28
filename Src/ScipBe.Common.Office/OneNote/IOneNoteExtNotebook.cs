using System.Collections.Generic;

namespace ScipBe.Common.Office.OneNote
{
    /// <summary>
    /// Notebook in OneNote with collection of Sections.
    /// </summary>
    public interface IOneNoteExtNotebook : IOneNoteNotebook
    {
        /// <summary>
        /// Collection of sections.
        /// </summary>
        IEnumerable<IOneNoteExtSection> Sections { get; }
    }

}