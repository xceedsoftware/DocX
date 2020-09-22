/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  /// <summary>
  /// All Document types are derived from DocumentElement. 
  /// This class contains properties which every element of a Document must contain.
  /// </summary>
  public abstract class DocumentElement
  {
    #region Private Members

    private PackagePart _mainPart;

    #endregion

    #region Public Properties

    /// <summary>
    /// This is the actual Xml that gives this element substance. 
    /// For example, a Paragraphs Xml might look something like the following
    /// <p>
    ///     <r>
    ///         <t>Hello World!</t>
    ///     </r>
    /// </p>
    /// </summary>
    public XElement Xml { get; set; } 

    public PackagePart PackagePart
    {
      get
      {
        return _mainPart;
      }
      set
      {
        _mainPart = value;
      }
    }

    #endregion

    #region Internal Properties

    /// <summary>
    /// This is a reference to the Document object that this element belongs to.
    /// Every Document element is connected to a document.
    /// </summary>
    internal Document Document { get; set; }

    #endregion

    #region Constructors

    /// <summary>
    /// Store both the document and xml so that they can be accessed by derived types.
    /// </summary>
    /// <param name="document">The document that this element belongs to.</param>
    /// <param name="xml">The Xml that gives this element substance</param>
    public DocumentElement( Document document, XElement xml )
    {
      this.Document = document;
      this.Xml = xml;
    }

    #endregion
  }

  /// <summary>
  /// This class provides functions for inserting new DocXElements before or after the current DocumentElement.
  /// Only certain DocXElements can support these functions without creating invalid documents, at the moment these are Paragraphs and Table.
  /// </summary>
  public abstract class InsertBeforeOrAfter : DocumentElement
  {
    #region Constructors

    public InsertBeforeOrAfter( Document document, XElement xml ) 
      : base( document, xml )
    {
    }

    #endregion

    #region Public Methods

    public virtual void InsertPageBreakBeforeSelf()
    {
      XElement p = new XElement
      (
          XName.Get( "p", Document.w.NamespaceName ),
              new XElement
              (
                  XName.Get( "r", Document.w.NamespaceName ),
                      new XElement
                      (
                          XName.Get( "br", Document.w.NamespaceName ),
                          new XAttribute( XName.Get( "type", Document.w.NamespaceName ), "page" )
                      )
              )
      );

      Xml.AddBeforeSelf( p );
    }

    public virtual void InsertPageBreakAfterSelf()
    {
      XElement p = new XElement
      (
          XName.Get( "p", Document.w.NamespaceName ),
              new XElement
              (
                  XName.Get( "r", Document.w.NamespaceName ),
                      new XElement
                      (
                          XName.Get( "br", Document.w.NamespaceName ),
                          new XAttribute( XName.Get( "type", Document.w.NamespaceName ), "page" )
                      )
              )
      );

      Xml.AddAfterSelf( p );
    }

    public virtual Paragraph InsertCaptionAfterSelf( string captionText )
    {
      var p = this.InsertParagraphAfterSelf( captionText + " " );
      p.StyleId = "Caption";

      var fldSimple = new XElement( XName.Get( "fldSimple", Document.w.NamespaceName ) );
      fldSimple.Add( new XAttribute( XName.Get( "instr", Document.w.NamespaceName ), @" SEQ " + captionText + @" \* ARABIC " ) );

      var actualCaptions = this.Document.Xml.Descendants( XName.Get( "fldSimple", Document.w.NamespaceName ) )
                                            .Where( field => (field != null) 
                                                && (field.GetAttribute( XName.Get( "instr", Document.w.NamespaceName ) ) != null )
                                                && field.GetAttribute( XName.Get( "instr", Document.w.NamespaceName ) ).StartsWith(" SEQ " + captionText) );
      var captionNumber = actualCaptions.Count() + 1;

      var content = XElement.Parse( string.Format(
       @"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
           <w:rPr>
              <w:noProof /> 
           </w:rPr>
           <w:t>{0}</w:t> 
         </w:r>",
       captionNumber)
      );
      fldSimple.Add( content );

      p.Xml.Add( fldSimple );

      return p;
    }

    public virtual Paragraph InsertParagraphBeforeSelf( Paragraph p )
    {
      Xml.AddBeforeSelf( p.Xml );
      XElement newlyInserted = Xml.ElementsBeforeSelf().First();

      p.Xml = newlyInserted;

      return p;
    }

    public virtual Paragraph InsertParagraphAfterSelf( Paragraph p )
    {
      Xml.AddAfterSelf( p.Xml );
      XElement newlyInserted = Xml.ElementsAfterSelf().First();

      //Dmitchern
      if( this as Paragraph != null )
        return new Paragraph( Document, newlyInserted, ( this as Paragraph )._endIndex );

      p.Xml = newlyInserted; //IMPORTANT: I think we have return new paragraph in any case, but I dont know what to put as startIndex parameter into Paragraph constructor
      return p;
    }

    public virtual Paragraph InsertParagraphBeforeSelf( string text )
    {
      return InsertParagraphBeforeSelf( text, false, new Formatting() );
    }

    public virtual Paragraph InsertParagraphAfterSelf( string text )
    {
      return InsertParagraphAfterSelf( text, false, new Formatting() );
    }

    public virtual Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges )
    {
      return InsertParagraphBeforeSelf( text, trackChanges, new Formatting() );
    }

    public virtual Paragraph InsertParagraphAfterSelf( string text, bool trackChanges )
    {
      return InsertParagraphAfterSelf( text, trackChanges, new Formatting() );
    }

    public virtual Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges, Formatting formatting )
    {
      XElement newParagraph = new XElement
      (
          XName.Get( "p", Document.w.NamespaceName ), new XElement( XName.Get( "pPr", Document.w.NamespaceName ) ), HelperFunctions.FormatInput( text, formatting.Xml )
      );

      if( trackChanges )
        newParagraph = Paragraph.CreateEdit( EditType.ins, DateTime.Now, newParagraph );

      Xml.AddBeforeSelf( newParagraph );
      XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

      return new Paragraph( Document, newlyInserted, -1 );
    }

    public virtual Paragraph InsertParagraphAfterSelf( string text, bool trackChanges, Formatting formatting )
    {
      XElement newParagraph = new XElement
      (
          XName.Get( "p", Document.w.NamespaceName ), new XElement( XName.Get( "pPr", Document.w.NamespaceName ) ), HelperFunctions.FormatInput( text, formatting.Xml )
      );

      if( trackChanges )
        newParagraph = Paragraph.CreateEdit( EditType.ins, DateTime.Now, newParagraph );

      Xml.AddAfterSelf( newParagraph );
      XElement newlyInserted = Xml.ElementsAfterSelf().First();

      Paragraph p = new Paragraph( Document, newlyInserted, -1 );

      return p;
    }

    public virtual Table InsertTableAfterSelf( int rowCount, int columnCount )
    {
      var newTable = HelperFunctions.CreateTable( rowCount, columnCount );
      Xml.AddAfterSelf( newTable );
      var newlyInserted = this.Xml.ElementsAfterSelf().First();

      var table = new Table( this.Document, newlyInserted, this.Document.PackagePart );
      table.PackagePart = this.PackagePart;
      return table;
    }

    public virtual Table InsertTableAfterSelf( Table t )
    {
      this.AddMissingPicturesInDocument( t );

      this.Xml.AddAfterSelf( t.Xml );
      var newlyInserted = this.Xml.ElementsAfterSelf().First();

      var table = new Table( this.Document, newlyInserted, this.Document.PackagePart );
      table.PackagePart = this.PackagePart;
      return table;
    }

    public virtual Table InsertTableBeforeSelf( int rowCount, int columnCount )
    {
      var newTable = HelperFunctions.CreateTable( rowCount, columnCount );
      this.Xml.AddBeforeSelf( newTable );
      var newlyInserted = this.Xml.ElementsBeforeSelf().Last();

      var table = new Table( this.Document, newlyInserted, this.Document.PackagePart );
      table.PackagePart = this.PackagePart;
      return table;
    }

    public virtual Table InsertTableBeforeSelf( Table t )
    {
      this.AddMissingPicturesInDocument( t );

      this.Xml.AddBeforeSelf( t.Xml );
      var newlyInserted = this.Xml.ElementsBeforeSelf().Last();

      var table = new Table( this.Document, newlyInserted, this.Document.PackagePart );
      table.PackagePart = this.PackagePart;
      return table;
    }

    public virtual List InsertListAfterSelf( List list )
    {
      for( var i = list.Items.Count - 1; i >= 0; --i )
      {
        this.Xml.AddAfterSelf( list.Items[ i ].Xml );
      }
      return list;
    }

    public virtual List InsertListBeforeSelf( List list )
    {
      foreach( var item in list.Items )
      {
        this.Xml.AddBeforeSelf( item.Xml );
      }
      return list;
    }

    #endregion

    #region Private Methods

    private void AddMissingPicturesInDocument( Table t )
    {
      if( t == null )
        return;

      // Make sure the pictures included in the Table are in the Document. If not, add them first.
      foreach( var p in t.Paragraphs )
      {
        if( p.Pictures.Count > 0 )
        {
          foreach( var pic in p.Pictures )
          {
            // Check if picture exists in Document.
            bool imageExists = false;
            foreach( var item in this.Document.PackagePart.GetRelationshipsByType( Document.RelationshipImage ) )
            {
              var targetUri = item.TargetUri.ToString();
              if( targetUri.Contains( pic.FileName ) )
              {
                imageExists = true;
                break;
              }
            }
            // Picture doesn't exists in Document, add it.
            if( !imageExists )
            {
              var newImage = this.Document.AddImage( pic.Stream );
              var newPicture = newImage.CreatePicture( pic.Height, pic.Width );
              p.PackagePart = this.Document.PackagePart;
              p.ReplacePicture( pic, newPicture );
            }
          }
        }
      }
    }

    #endregion
  }

  public static class XmlTemplates
  {
    #region Public Constants

    public const string TableOfContentsXmlBase = @"
        <w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
            <w:sdtPr>
            <w:docPartObj>
                <w:docPartGallery w:val='Table of Contents'/>
                <w:docPartUnique/>
            </w:docPartObj>\
            </w:sdtPr>
            <w:sdtEndPr>
            <w:rPr>
                <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>
                <w:color w:val='auto'/>
                <w:sz w:val='22'/>
                <w:szCs w:val='22'/>
                <w:lang w:eastAsia='en-US'/>
            </w:rPr>
            </w:sdtEndPr>
            <w:sdtContent>
            <w:p>
                <w:pPr>
                <w:pStyle w:val='{0}'/>
                </w:pPr>
                <w:r>
                <w:t>{1}</w:t>
                </w:r>
            </w:p>
            <w:p>
                <w:pPr>
                <w:pStyle w:val='TOC1'/>
                <w:tabs>
                    <w:tab w:val='right' w:leader='dot' w:pos='{2}'/>
                </w:tabs>
                <w:rPr>
                    <w:noProof/>
                </w:rPr>
                </w:pPr>
                <w:r>
                <w:fldChar w:fldCharType='begin' w:dirty='true'/>
                </w:r>
                <w:r>
                <w:instrText xml:space='preserve'> {3} </w:instrText>
                </w:r>
                <w:r>
                <w:fldChar w:fldCharType='separate'/>
                </w:r>
            </w:p>
            <w:p>
                <w:r>
                <w:rPr>
                    <w:b/>
                    <w:bCs/>
                    <w:noProof/>
                </w:rPr>
                <w:fldChar w:fldCharType='end'/>
                </w:r>
            </w:p>
            </w:sdtContent>
        </w:sdt>
        ";

    public const string TableOfContentsHeadingStyleBase = @"
        <w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
            <w:name w:val='TOC Heading'/>
            <w:basedOn w:val='Heading1'/>
            <w:next w:val='Normal'/>
            <w:uiPriority w:val='39'/>
            <w:semiHidden/>
            <w:unhideWhenUsed/>
            <w:qFormat/>
            <w:rsid w:val='00E67AA6'/>
            <w:pPr>
            <w:outlineLvl w:val='9'/>
            </w:pPr>
            <w:rPr>
            <w:lang w:eastAsia='nb-NO'/>
            </w:rPr>
        </w:style>
        ";

    internal const int TableOfContentsElementDefaultIndentation = 220;

    public const string TableOfContentsElementStyleBase = @"  
        <w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
            <w:name w:val='{1}' />
            <w:basedOn w:val='Normal' />
            <w:next w:val='Normal' />
            <w:autoRedefine />
            <w:uiPriority w:val='39' />
            <w:unhideWhenUsed />
            <w:pPr>
            <w:spacing w:after='100' />
            <w:ind w:left='{2}' />
            </w:pPr>
        </w:style>
        ";

    public const string TableOfContentsHyperLinkStyleBase = @"  
        <w:style w:type='character' w:styleId='Hyperlink' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
            <w:name w:val='Hyperlink' />
            <w:basedOn w:val='Normal' />
            <w:uiPriority w:val='99' />
            <w:unhideWhenUsed />
            <w:rPr>
            <w:color w:val='0000FF' w:themeColor='hyperlink' />
            <w:u w:val='single' />
            </w:rPr>
        </w:style>
        ";

    #endregion
  }
}
