using System.Drawing;

namespace ScipBe.Common.Office.OneNote
{
    /// <summary>
    /// Section in OneNote.
    /// </summary>
    public interface IOneNoteSection
    {
        /// <summary>
        /// ID of Section.
        /// </summary>
        string ID { get; }

        /// <summary>
        /// Name of Section.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Physical file path of Section.
        /// </summary>
        string Path { get; }

        /// <summary>
        /// Is Section encrypted?
        /// </summary>
        bool Encrypted { get; }

        /// <summary>
        /// Color of tab of section.
        /// </summary>
        Color? Color { get; }
    }
}