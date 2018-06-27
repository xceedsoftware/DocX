/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Collections.ObjectModel;

namespace Xceed.Words.NET
{
  public class Footer : Container, IParagraphContainer
  {
    #region Public Properties

    public bool PageNumbers
    {
      get
      {
        return false;
      }

      set
      {
        XElement e = XElement.Parse
        ( @"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                    <w:sdtPr>
                      <w:id w:val='157571950' />
                      <w:docPartObj>
                        <w:docPartGallery w:val='Page Numbers (Top of Page)' />
                        <w:docPartUnique />
                      </w:docPartObj>
                    </w:sdtPr>
                    <w:sdtContent>
                      <w:p w:rsidR='008D2BFB' w:rsidRDefault='008D2BFB'>
                        <w:pPr>
                          <w:pStyle w:val='Header' />
                          <w:jc w:val='center' />
                        </w:pPr>
                        <w:fldSimple w:instr=' PAGE \* MERGEFORMAT'>
                          <w:r>
                            <w:rPr>
                              <w:noProof />
                            </w:rPr>
                            <w:t>1</w:t>
                          </w:r>
                        </w:fldSimple>
                      </w:p>
                    </w:sdtContent>
                  </w:sdt>"
       );

        Xml.AddFirst( e );
      }
    }

    public override ReadOnlyCollection<Paragraph> Paragraphs
    {
      get
      {
        var paragraphs = base.Paragraphs;
        foreach( var paragraph in paragraphs )
        {
          paragraph.PackagePart = this.PackagePart;
        }
        return paragraphs;
      }
    }

    public override List<Table> Tables
    {
      get
      {
        var l = base.Tables;
        l.ForEach( x => x.PackagePart = this.PackagePart );
        return l;
      }
    }

    public List<Image> Images
    {
      get
      {
        var imageRelationships = this.PackagePart.GetRelationshipsByType( DocX.RelationshipImage );
        if( imageRelationships.Count() > 0 )
        {
          return
          (
              from i in imageRelationships
              select new Image( Document, i )
          ).ToList();
        }

        return new List<Image>();
      }
    }

    #endregion

    #region Constructors

    internal Footer( DocX document, XElement xml, PackagePart mainPart ) : base( document, xml )
    {
      this.PackagePart = mainPart;
    }

    #endregion

    #region Public Methods

    public override Paragraph InsertParagraph()
    {
      var p = base.InsertParagraph();
      p.PackagePart = this.PackagePart;
      return p;
    }

    public override Paragraph InsertParagraph( int index, string text, bool trackChanges )
    {
      var p = base.InsertParagraph( index, text, trackChanges );
      p.PackagePart = this.PackagePart;
      return p;
    }

    public override Paragraph InsertParagraph( Paragraph p )
    {
      p.PackagePart = this.PackagePart;
      return base.InsertParagraph( p );
    }

    public override Paragraph InsertParagraph( int index, Paragraph p )
    {
      p.PackagePart = this.PackagePart;
      return base.InsertParagraph( index, p );
    }

    public override Paragraph InsertParagraph( int index, string text, bool trackChanges, Formatting formatting )
    {
      var p = base.InsertParagraph( index, text, trackChanges, formatting );
      p.PackagePart = this.PackagePart;
      return p;
    }

    public override Paragraph InsertParagraph( string text )
    {
      var p = base.InsertParagraph( text );
      p.PackagePart = this.PackagePart;
      return p;
    }

    public override Paragraph InsertParagraph( string text, bool trackChanges )
    {
      var p = base.InsertParagraph( text, trackChanges );
      p.PackagePart = this.PackagePart;
      return p;
    }

    public override Paragraph InsertParagraph( string text, bool trackChanges, Formatting formatting )
    {
      var p = base.InsertParagraph( text, trackChanges, formatting );
      p.PackagePart = this.PackagePart;

      return p;
    }

    public override Paragraph InsertEquation( String equation )
    {
      var p = base.InsertEquation( equation );
      p.PackagePart = this.PackagePart;
      return p;
    }

    public override Table InsertTable( int rowCount, int columnCount )
    {
      var table =  base.InsertTable( rowCount, columnCount );
      return this.SetMainPart( table );
    }

    public override Table InsertTable( int index, Table t )
    {
      var table = base.InsertTable( index, t );
      return this.SetMainPart( table );
    }

    public override Table InsertTable( Table t )
    {
      var table = base.InsertTable( t );
      return this.SetMainPart( table );
    }

    public override Table InsertTable( int index, int rowCount, int columnCount )
    {
      var table = base.InsertTable( index, rowCount, columnCount );
      return this.SetMainPart( table );
    }

    #endregion

    #region Private Methods

    private Table SetMainPart( Table table )
    {
      if( table != null )
      {
        table.PackagePart = this.PackagePart;
      }
      return table;
    }

    #endregion
  }
}
