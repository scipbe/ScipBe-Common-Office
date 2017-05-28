// ==============================================================================================
// Namespace   : ScipBe.Common.Office
// Author      : Stefan Cruysberghs
// Website     : http://www.scip.be
// Status      : Open source - MIT License
// ==============================================================================================

using System;
using System.Collections.Generic;
using System.Drawing;

namespace ScipBe.Common.Office
{
  /// <summary>
  /// Notebook in OneNote
  /// </summary>
  public interface IOneNoteNotebook
  {
    /// <summary>
    /// ID of Notebook
    /// </summary>
    string ID { get; }
    /// <summary>
    /// Name of Notebook
    /// </summary>
    string Name { get; }
    /// <summary>
    /// Nickname of Notebook
    /// </summary>
    string NickName { get; }
    /// <summary>
    /// Physical file path of Notebook
    /// </summary>
    string Path { get; }
    /// <summary>
    /// Color of tab of Notebook
    /// </summary>
    Color Color { get; }
  }

   /// <summary>
   /// Notebook in OneNote with collection of Sections
   /// </summary>
   public interface IOneNoteExtNotebook : IOneNoteNotebook
   {
     /// <summary>
     /// Collection of sections
     /// </summary>
     IEnumerable<IOneNoteExtSection> Sections { get; }
   }

   /// <summary>
   /// Section in OneNote
   /// </summary>
   public interface IOneNoteSection
   {
     /// <summary>
     /// ID of Section
     /// </summary>
     string ID { get; }
     /// <summary>
     /// Name of Section
     /// </summary>
     string Name { get; }
     /// <summary>
     /// Physical file path of Section
     /// </summary>
     string Path { get; }
     /// <summary>
     /// Is Section encrypted?
     /// </summary>
     bool Encrypted { get; }
     /// <summary>
     /// Color of tab of section
     /// </summary>
     Color Color { get; }
   }

  /// <summary>
  /// Section in OneNote with collection of Pages
  /// </summary>
   public interface IOneNoteExtSection : IOneNoteSection
   {
     /// <summary>
     /// Collection of pages
     /// </summary>
     IEnumerable<IOneNotePage> Pages { get; }
   }

   /// <summary>
   /// Page in OneNote
   /// </summary>
   public interface IOneNotePage
   {
     /// <summary>
     /// ID of Page
     /// </summary>
     string ID { get; }
     /// <summary>
     /// Name of Page
     /// </summary>
     string Name { get; }
     /// <summary>
     /// Date and time of creation of the Page
     /// </summary>
     DateTime DateTime { get; }
     /// <summary>
     /// Date and time of last modification of Page
     /// </summary>
     DateTime LastModified { get; }
   }

   /// <summary>
   /// Page with reference to Section and Notebook
   /// </summary>
   public interface IOneNoteExtPage : IOneNotePage
   {
     /// <summary>
     /// Section of Page
     /// </summary>
     IOneNoteSection Section { get; }
     /// <summary>
     /// Notebook of Page
     /// </summary>
     IOneNoteNotebook Notebook { get; }
   }

}