using System.Drawing;

namespace ScipBe.Common.Office.OneNote
{
    /// <summary>
    /// Notebook in OneNote.
    /// </summary>
    public interface IOneNoteNotebook
    {
        /// <summary>
        /// ID of Notebook.
        /// </summary>
        string ID { get; }
        /// <summary>
        /// Name of Notebook.
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Nickname of Notebook.
        /// </summary>
        string NickName { get; }
        /// <summary>
        /// Physical file path of Notebook.
        /// </summary>
        string Path { get; }
        /// <summary>
        /// Color of tab of Notebook.
        /// </summary>
        Color? Color { get; }
    }
}