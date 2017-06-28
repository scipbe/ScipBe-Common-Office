using System.Collections.Generic;
using System.Drawing;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNoteExtSection : IOneNoteExtSection
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public bool Encrypted { get; set; }
        public Color Color { get; set; }
        public IEnumerable<IOneNotePage> Pages { get; set; }
    }
}