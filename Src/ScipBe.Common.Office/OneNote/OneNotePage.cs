using System;
using Microsoft.Office.Interop.OneNote;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNotePage : IOneNotePage
    {
        private readonly Application oneNote;

        public OneNotePage(Application oneNote)
        {
            this.oneNote = oneNote;
        }

        public string ID { get; set; }
        public string Name { get; set; }
        public int Level { get; set; }
        public DateTime DateTime { get; set; }
        public DateTime LastModified { get; set; }

        public void OpenInOneNote()
        {
            this.oneNote.NavigateTo(this.ID);
        }
    }
}