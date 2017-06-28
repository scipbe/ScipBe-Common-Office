using System;

namespace ScipBe.Common.Office.OneNote
{
    internal class OneNotePage : IOneNotePage
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public int Level { get; set; }
        public DateTime DateTime { get; set; }
        public DateTime LastModified { get; set; }
    }
}