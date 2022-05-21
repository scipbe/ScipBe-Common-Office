using System.Collections.Generic;
using System.Drawing;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNoteExtSection : OneNoteSection, IOneNoteExtSection
    {
        public IEnumerable<IOneNotePage> Pages { get; set; }
    }
}