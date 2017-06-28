using System;

namespace ScipBe.Common.Office.OneNote
{
    /// <summary>
    /// Page in OneNote.
    /// </summary>
    public interface IOneNotePage
    {
        /// <summary>
        /// ID of Page.
        /// </summary>
        string ID { get; }
        /// <summary>
        /// Name of Page.
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Level of page.
        /// </summary>
        int Level { get; }
        /// <summary>
        /// Date and time of creation of the Page.
        /// </summary>
        DateTime DateTime { get; }
        /// <summary>
        /// Date and time of last modification of Page.
        /// </summary>
        DateTime LastModified { get; }
    }
}