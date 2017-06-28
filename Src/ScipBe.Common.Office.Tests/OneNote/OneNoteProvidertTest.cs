using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using ScipBe.Common.Office.OneNote;

namespace ScipBe.Common.Office.Tests
{
    [TestClass]
    public class OneNoteProvidertTest
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
    }
}
