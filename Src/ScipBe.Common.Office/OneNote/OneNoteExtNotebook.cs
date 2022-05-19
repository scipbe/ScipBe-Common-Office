using System.Collections.Generic;
using System.Drawing;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNoteExtNotebook : OneNoteNotebook, IOneNoteExtNotebook
    {
        public IEnumerable<IOneNoteExtSection> Sections { get; set; }
    }
}