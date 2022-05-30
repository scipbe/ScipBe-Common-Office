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
            Assert.IsTrue(notebooks.Any());
        }

        [TestMethod]
        public void Pages()
        {
            // Act
            var pages = OneNoteProvider.PageItems;

            // Arrange
            Assert.IsTrue(pages.Any());
        }

        [TestMethod]
        public void EnumerateSections()
        {
            // Act
            var sections = OneNoteProvider.NotebookItems.SelectMany(n => n.Sections);

            // Arrange
            Assert.IsTrue(sections.Any());
        }

        [TestMethod]
        public void FindPages()
        {
            // Act
            var pages = OneNoteProvider.FindPages("the");

            // Arrange
            Assert.IsTrue(pages.Any());
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
    }
}
