using System.Linq;
using ScipBe.Common.Office.Outlook;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ScipBe.Common.Office.Tests
{
    [TestClass]
    public class OutlookProvidertTest
    {
        [TestMethod]
        public void Folders()
        {
            // Arrange
            var outlook = new OutlookProvider();

            // Act
            var folders = outlook.Folders;

            // Assert
            Assert.IsTrue(outlook.Folders.Count() > 0);
        }

        [TestMethod]
        public void ContactItems()
        {
            // Arrange
            var outlook = new OutlookProvider();

            // Act
            var contactItems = outlook.ContactItems;

            // Assert
            Assert.IsTrue(contactItems.Count() > 0);
        }

        [TestMethod]
        public void CalendarItems()
        {
            // Arrange
            var outlook = new OutlookProvider();

            // Act
            var calendarItems = outlook.CalendarItems;

            // Assert
            Assert.IsTrue(calendarItems.Count() > 0);
        }

        [TestMethod]
        public void InboxItems()
        {
            // Arrange
            var outlook = new OutlookProvider();

            // Act
            var inboxItems = outlook.InboxItems;

            // Assert
            Assert.IsTrue(inboxItems.Count() > 0);
        }

        [TestMethod]
        public void SentMailItems()
        {
            // Arrange
            var outlook = new OutlookProvider();

            // Act
            var sentMailItems = outlook.SentMailItems;

            // Assert
            Assert.IsTrue(sentMailItems.Count() > 0);
        }

        [TestMethod]
        public void GetItemsByFolder()
        {
            // Arrange
            var outlook = new OutlookProvider();

            // Act
            var folder = outlook.Folders.FirstOrDefault(f => f.Items.Count > 0 && f.DefaultItemType == OlItemType.olMailItem);
            var mailItems = outlook.GetItems<MailItem>(folder);

            // Assert
            Assert.IsTrue(mailItems.Count() > 0);
        }

        [TestMethod]
        public void GetItemByFolderPath()
        {
            // Arrange
            var outlook = new OutlookProvider();

            // Act
            var folder = outlook.Folders.FirstOrDefault(f => f.Items.Count > 0 && f.DefaultItemType == OlItemType.olAppointmentItem);
            var appointmentItems = outlook.GetItems<AppointmentItem>(folder.FolderPath);

            // Assert
            Assert.IsTrue(appointmentItems.Count() > 0);
        }
    }
}
