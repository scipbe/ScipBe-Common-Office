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
            // Arrange
            var oneNoteProvider = new OneNoteProvider();

            // Act
            var notebooks = oneNoteProvider.NotebookItems;

            // Arrange
            Assert.IsTrue(notebooks.Count() > 0);
        }

        [TestMethod]
        public void Pages()
        {
            // Arrange
            var oneNoteProvider = new OneNoteProvider();

            // Act
            var pages = oneNoteProvider.PageItems;

            // Arrange
            Assert.IsTrue(pages.Count() > 0);
        }

        [TestMethod]
        public void EnumerateSections()
        {
            // Arrange
            var oneNoteProvider = new OneNoteProvider();

            // Act
            var sections = oneNoteProvider.NotebookItems.SelectMany(n => n.Sections);

            // Arrange
            Assert.IsTrue(sections.Any());
        }

        [TestMethod]
        public void FindPages()
        {
            // Arrange
            var oneNoteProvider = new OneNoteProvider();

            // Act
            var pages = oneNoteProvider.FindPages("the");

            // Arrange
            Assert.IsTrue(pages.Any());
        }
    }
}
