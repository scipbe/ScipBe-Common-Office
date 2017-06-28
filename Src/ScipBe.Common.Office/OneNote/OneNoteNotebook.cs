using System.Drawing;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNoteNotebook : IOneNoteNotebook
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string NickName { get; set; }
        public string Path { get; set; }
        public Color Color { get; set; }
    }
}