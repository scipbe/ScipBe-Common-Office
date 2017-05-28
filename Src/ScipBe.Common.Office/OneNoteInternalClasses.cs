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
  internal class OneNoteNotebook : IOneNoteNotebook
  {
    public string ID { get; set; }
    public string Name { get; set; }
    public string NickName { get; set; }
    public string Path { get; set; }
    public Color Color { get; set; }
  }

  internal class OneNoteExtNotebook : IOneNoteExtNotebook
  {
    public string ID { get; set; }
    public string Name { get; set; }
    public string NickName { get; set; }
    public string Path { get; set; }
    public Color Color { get; set; }
    public IEnumerable<IOneNoteExtSection> Sections { get; set; }
  }

  internal class OneNoteSection : IOneNoteSection
  {
    public string ID { get; set; }
    public string Name { get; set; }
    public string Path { get; set; }
    public bool Encrypted { get; set; }
    public Color Color { get; set; }
  }

  internal class OneNoteExtSection : IOneNoteExtSection
  {
    public string ID { get; set; }
    public string Name { get; set; }
    public string Path { get; set; }
    public bool Encrypted { get; set; }
    public Color Color { get; set; }    
    public IEnumerable<IOneNotePage> Pages { get; set; }
  }

  internal class OneNotePage : IOneNotePage
  {
    public string ID { get; set; }
    public string Name { get; set; }
    public DateTime DateTime { get; set; }
    public DateTime LastModified { get; set; }
  }

  internal class OneNoteExtPage : IOneNoteExtPage
  {
    public string ID { get; set; }
    public string Name { get; set; }
    public DateTime DateTime { get; set; }
    public DateTime LastModified { get; set; }
    public IOneNoteSection Section { get; set; }
    public IOneNoteNotebook Notebook { get; set; }
  }
}