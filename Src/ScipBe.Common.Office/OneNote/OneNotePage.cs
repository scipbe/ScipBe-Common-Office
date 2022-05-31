using System;
using Microsoft.Office.Interop.OneNote;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNotePage : IOneNotePage
    {
        public OneNotePage()
        {
        }

        public string ID { get; set; }
        public string Name { get; set; }
        public int Level { get; set; }
        public DateTime DateTime { get; set; }
        public DateTime LastModified { get; set; }

        public string GetContent()
        {
            return OneNoteProvider.CallOneNoteSafely(oneNote =>
            {
                oneNote.GetPageContent(this.ID, out string content);
                return content;
            });
        }

        public void OpenInOneNote()
        {
            OneNoteProvider.CallOneNoteSafely(oneNote =>
            {
                oneNote.NavigateTo(this.ID);
                return true;
            });
        }
    }
}