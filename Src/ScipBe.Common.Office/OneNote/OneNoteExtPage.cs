using Microsoft.Office.Interop.OneNote;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNoteExtPage : OneNotePage, IOneNoteExtPage
    {
        public OneNoteExtPage(Application oneNote)
            : base(oneNote) { }

        public IOneNoteSection Section { get; set; }
        public IOneNoteNotebook Notebook { get; set; }
    }
}