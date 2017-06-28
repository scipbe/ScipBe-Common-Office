namespace ScipBe.Common.Office.OneNote
{

    /// <summary>
    /// Page with reference to Section and Notebook.
    /// </summary>
    public interface IOneNoteExtPage : IOneNotePage
    {
        /// <summary>
        /// Section of Page.
        /// </summary>
        IOneNoteSection Section { get; }
        /// <summary>
        /// Notebook of Page.
        /// </summary>
        IOneNoteNotebook Notebook { get; }
    }
}