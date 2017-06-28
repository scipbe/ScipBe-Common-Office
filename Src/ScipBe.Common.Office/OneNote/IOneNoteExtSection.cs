using System.Collections.Generic;

namespace ScipBe.Common.Office.OneNote
{
    /// <summary>
    /// Section in OneNote with collection of Pages.
    /// </summary>
    public interface IOneNoteExtSection : IOneNoteSection
    {
        /// <summary>
        /// Collection of pages.
        /// </summary>
        IEnumerable<IOneNotePage> Pages { get; }
    }
}