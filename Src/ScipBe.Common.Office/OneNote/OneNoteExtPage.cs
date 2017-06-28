using System;
using System.Collections.Generic;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNoteExtPage : IOneNoteExtPage
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public int Level { get; set; }
        public DateTime DateTime { get; set; }
        public DateTime LastModified { get; set; }
        public IOneNoteSection Section { get; set; }
        public IOneNoteNotebook Notebook { get; set; }
    }
}