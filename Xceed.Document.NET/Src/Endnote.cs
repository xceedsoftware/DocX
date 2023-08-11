/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2023 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.IO.Packaging;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public class Endnote: Note
  {
    #region Constructor

    internal Endnote( Document document, PackagePart part, XElement xml ) : base( document, part, xml )
    {
    }

    #endregion

    #region Overrides

    internal override string GetNoteRefType()
    {
      return "endnoteRef";
    }

    internal override XElement CreateReferenceRunCore( bool customMarkFollows, XElement symbol, Formatting noteNumberFormatting )
    {
      var rPr = (noteNumberFormatting != null)
                ? noteNumberFormatting.Xml 
                : new XElement( XName.Get( "rPr", Document.w.NamespaceName ) );

      rPr.AddFirst( new XElement( XName.Get( "rStyle", Document.w.NamespaceName ), 
                                            new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "EndnoteReference" ) ) );

      var r = new XElement( XName.Get( "r", Document.w.NamespaceName ) );
      r.Add( rPr );

      var endNoteReference = new XElement( XName.Get( "endnoteReference", Document.w.NamespaceName ),
                                                          new XAttribute( XName.Get( "id", Document.w.NamespaceName ), this.Id ) );
      if( customMarkFollows )
      {
        endNoteReference.SetAttributeValue( XName.Get( "customMarkFollows", Document.w.NamespaceName ), "1" );
      }
      if( symbol != null )
      {
        endNoteReference.Add( symbol );
      }

      r.Add( endNoteReference );

      return r;
    }

    #endregion
  }
}
