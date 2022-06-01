/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2022 Xceed Software Inc.
 
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
  public class Footnote : Note
  {
    #region Constructor

    internal Footnote( Document document, PackagePart part, XElement xml ) : base( document, part, xml )
    {
    }

    #endregion

    #region Overrides

    internal override string GetNoteRefType()
    {
      return "footnoteRef";
    }

    internal override XElement CreateReferenceRunCore( bool customMarkFollows, XElement symbol, Formatting noteNumberFormatting )
    {
      var rPr = ( noteNumberFormatting != null )
                 ? noteNumberFormatting.Xml
                 : new XElement( XName.Get( "rPr", Document.w.NamespaceName ) );

      rPr.Add( new XElement( XName.Get( "rStyle", Document.w.NamespaceName ),
                                            new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "FootnoteReference" ) ) );

      var r = new XElement( XName.Get( "r", Document.w.NamespaceName ) );
      r.Add( rPr );

      var footNoteReference = new XElement( XName.Get( "footnoteReference", Document.w.NamespaceName ),
                                                          new XAttribute( XName.Get( "id", Document.w.NamespaceName ), this.Id ) );
      if( customMarkFollows )
      {
        footNoteReference.SetAttributeValue( XName.Get( "customMarkFollows", Document.w.NamespaceName ), "1" );
      }
      if( symbol != null )
      {
        footNoteReference.Add( symbol );
      }

      r.Add( footNoteReference );

      return r;
    }

    #endregion
  }
}
