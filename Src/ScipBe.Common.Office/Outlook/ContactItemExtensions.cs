using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.IO;

namespace ScipBe.Common.Office.Outlook
{
    public static class ContactItemExtensions
    {

        /// <summary>
        /// Save contact picture as JPG in given folder and return path to this file.
        /// </summary>
        /// <param name="contact">Outlook contact item.</param>
        /// <param name="path">Path to folder with temporary Outlook contact pictures.</param>
        /// <returns>Path to JPG file with Outlook contact picture.</returns>
        public static string GetContactPicturePath(this ContactItem contact, string path)
        {
            string picturePath = "";

            if (contact.HasPicture)
            {
                foreach (Attachment att in contact.Attachments)
                {
                    if (att.DisplayName == "ContactPicture.jpg")
                    {
                        try
                        {
                            picturePath = Path.GetDirectoryName(path) + "\\Contact_" + contact.EntryID + ".jpg";
                            if (!File.Exists(picturePath))
                            {
                                att.SaveAsFile(picturePath);
                            }
                        }
                        catch (IOException ex)
                        {
                            picturePath = "";
                            Debug.Write(ex.Message);
                        }
                    }
                }
            }

            return picturePath;
        }

        /// <summary>
        /// Save contact picture as JPG in Windows temporary folder and return path to this file.
        /// </summary>
        /// <param name="contact">Outlook contact item.</param>
        /// <returns>Path to JPG file with Outlook contact picture.</returns>
        public static string GetContactPicturePath(ContactItem contact)
        {
            return contact.GetContactPicturePath(Path.GetTempPath());
        }
    }
}
