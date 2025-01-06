using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using ScipBe.Common.Office.OneNote;

namespace ScipBe.Common.Office.Tests
{
    [TestClass]
    public class OneNoteProviderTest
    {
        [TestMethod]
        public void Notebooks()
        {
            // Act
            var notebooks = OneNoteProvider.NotebookItems;

            // Arrange
            Assert.IsTrue(notebooks.Count() > 0);
        }

        [TestMethod]
        public void Pages()
        {
            // Act
            var pages = OneNoteProvider.PageItems;

            // Arrange
            Assert.IsTrue(pages.Count() > 0);
        }

        [TestMethod]
        public void EnumerateSections()
        {
            // Act
            var sections = OneNoteProvider.NotebookItems.SelectMany(n => n.Sections);

            // Arrange
            Assert.IsTrue(sections.Count() > 0);
        }

        [TestMethod]
        public void FindPages()
        {
            // Act
            var pages = OneNoteProvider.FindPages("the");

            // Arrange
            Assert.IsTrue(pages.Count() > 0);
        }

        [TestMethod]
        public void GetContent()
        {
            // Act
            var page = OneNoteProvider.PageItems.First();
            string content = page.GetContent();

            // Arrange
            Assert.IsFalse(string.IsNullOrWhiteSpace(content));
        }

        [TestMethod]
        public void OpenInOneNote()
        {
            // Act
            var page = OneNoteProvider.PageItems.First();
            page.OpenInOneNote();
        }
    }
}
