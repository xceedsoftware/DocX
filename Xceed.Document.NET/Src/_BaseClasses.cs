/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
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
  public abstract class DocumentElement
  {
    #region Private Members

    private PackagePart _mainPart;

    #endregion

    #region Public Properties

    public XElement Xml
    {
      get; set;
    }

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

    internal Document Document
    {
      get; set;
    }

    #endregion

    #region Constructors

    public DocumentElement( Document document, XElement xml )
    {
      this.Document = document;
      this.Xml = xml;
    }

    #endregion

    #region Internal Methods

    internal double GetAvailableWidth()
    {
      var currentSection = ( ( this.Document.Sections != null ) && ( this.Document.Sections.Count > 0 ) )
                          ? this.Document.Sections[ this.Document.Sections.Count - 1 ]
                          : null;
      if( currentSection != null )
        return Convert.ToDouble( currentSection.PageWidth - currentSection.MarginLeft - currentSection.MarginRight );

      return Convert.ToDouble( this.Document.PageWidth - this.Document.MarginLeft - this.Document.MarginRight );
    }

    #endregion
  }

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
                                            .Where( field => ( field != null )
                                                && ( field.GetAttribute( XName.Get( "instr", Document.w.NamespaceName ) ) != null )
                                                && field.GetAttribute( XName.Get( "instr", Document.w.NamespaceName ) ).StartsWith( " SEQ " + captionText ) );
      var actualCaptions2 = this.Document.Xml.Descendants( XName.Get( "instrText", Document.w.NamespaceName ) )
                                             .Where( field => ( field != null )
                                                  && field.Value.StartsWith( " SEQ " + captionText ) );
      var captionNumber = actualCaptions.Count() + actualCaptions2.Count() + 1;

      var content = XElement.Parse( string.Format(
       @"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
           <w:rPr>
              <w:noProof /> 
           </w:rPr>
           <w:t>{0}</w:t> 
         </w:r>",
       captionNumber )
      );
      fldSimple.Add( content );

      p.Xml.Add( fldSimple );

      return p;
    }

    public virtual Paragraph InsertParagraphBeforeSelf( Paragraph p )
    {
      this.Xml.AddBeforeSelf( p.Xml );

      var newlyInserted = this.Xml.ElementsBeforeSelf().Last();

      var newParagraph = new Paragraph( this.Document, newlyInserted, -1 );
      newParagraph.PackagePart = this.PackagePart;

      this.ClearMainParentContainerCache( p );

      return newParagraph;
    }

    public virtual Paragraph InsertParagraphAfterSelf( Paragraph p )
    {
      this.Xml.AddAfterSelf( p.Xml );

      var newlyInserted = this.Xml.ElementsAfterSelf().First();

      var newParagraph = new Paragraph( this.Document, newlyInserted, -1 );
      newParagraph.PackagePart = this.PackagePart;

      this.ClearMainParentContainerCache( p );

      return newParagraph;
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
      {
        newParagraph = Document.CreateEdit( EditType.ins, DateTime.Now, newParagraph );
      }

      this.Xml.AddBeforeSelf( newParagraph );
      XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

      var p = new Paragraph( this.Document, newlyInserted, -1 );
      p.PackagePart = this.PackagePart;

      this.ClearMainParentContainerCache( p );

      return p;
    }

    public virtual Paragraph InsertParagraphAfterSelf( string text, bool trackChanges, Formatting formatting )
    {
      XElement newParagraph = new XElement
      (
          XName.Get( "p", Document.w.NamespaceName ), new XElement( XName.Get( "pPr", Document.w.NamespaceName ) ), HelperFunctions.FormatInput( text, formatting.Xml )
      );

      if( trackChanges )
        newParagraph = Document.CreateEdit( EditType.ins, DateTime.Now, newParagraph );

      Xml.AddAfterSelf( newParagraph );
      XElement newlyInserted = Xml.ElementsAfterSelf().First();

      var p = new Paragraph( this.Document, newlyInserted, -1 );
      p.PackagePart = this.PackagePart;

      this.ClearMainParentContainerCache( p );

      return p;
    }

    public virtual Table InsertTableAfterSelf( int rowCount, int columnCount )
    {
      var newTableXElement = HelperFunctions.CreateTable( rowCount, columnCount, this.GetAvailableWidth() );
      this.Xml.AddAfterSelf( newTableXElement );
      var newlyInsertedXElement = this.Xml.ElementsAfterSelf().First();

      var table = new Table( this.Document, newlyInsertedXElement, this.Document.PackagePart );
      table.PackagePart = this.PackagePart;

      foreach( var p in table.Paragraphs )
      {
        this.AddParagraphInCache( p );
      }

      return table;
    }

    public virtual Table InsertTableAfterSelf( Table t )
    {
      this.AddMissingPicturesInDocument( t );

      // Use a copy to be able to insert the table in header/footer/document.
      var tableCopyXml = new XElement( t.Xml );

      this.Xml.AddAfterSelf( tableCopyXml );

      var newlyInserted = this.Xml.ElementsAfterSelf().First();
      var newTable = new Table( this.Document, newlyInserted, this.Document.PackagePart );
      newTable.PackagePart = this.PackagePart;

      this.AddPicturesInPackage( newTable );

      foreach( var p in newTable.Paragraphs )
      {
        this.AddParagraphInCache( p );
      }

      return newTable;
    }

    public virtual Table InsertTableBeforeSelf( int rowCount, int columnCount )
    {
      var newTableXElement = HelperFunctions.CreateTable( rowCount, columnCount, this.GetAvailableWidth() );
      this.Xml.AddBeforeSelf( newTableXElement );
      var newlyInsertedXElement = this.Xml.ElementsBeforeSelf().Last();

      var table = new Table( this.Document, newlyInsertedXElement, this.Document.PackagePart );
      table.PackagePart = this.PackagePart;

      foreach( var p in table.Paragraphs )
      {
        this.AddParagraphInCache( p );
      }
      return table;
    }

    public virtual Table InsertTableBeforeSelf( Table t )
    {
      this.AddMissingPicturesInDocument( t );

      // Use a copy to be able to insert the table in header/footer/document.
      var tableCopyXml = new XElement( t.Xml );

      this.Xml.AddBeforeSelf( tableCopyXml );

      var newlyInserted = this.Xml.ElementsBeforeSelf().Last();
      var newTable = new Table( this.Document, newlyInserted, this.Document.PackagePart );
      newTable.PackagePart = this.PackagePart;

      this.AddPicturesInPackage( newTable );

      foreach( var p in newTable.Paragraphs )
      {
        this.AddParagraphInCache( p );
      }

      return newTable;
    }

    public virtual List InsertListAfterSelf( List list )
    {
      for( var i = list.Items.Count - 1; i >= 0; --i )
      {
        var listItem = list.Items[ i ];
        this.Xml.AddAfterSelf( listItem.Xml );
      }

      if( list.Items.Count > 0 )
      {
        var mainParentContainer = list.Items.LastOrDefault().GetMainParentContainer();
        if( mainParentContainer != null )
        {
          mainParentContainer.ClearParagraphsCache();
        }
      }

      this.Document.ClearParagraphsCache();
      return list;
    }

    public virtual List InsertListBeforeSelf( List list )
    {
      for( int i = 0; i < list.Items.Count; i++ )
      {
        var item = list.Items[ i ];
        this.Xml.AddBeforeSelf( item.Xml );
      }

      if( list.Items.Count > 0 )
      {
        var mainParentContainer = list.Items.LastOrDefault().GetMainParentContainer();
        if( mainParentContainer != null )
        {
          mainParentContainer.ClearParagraphsCache();
        }
      }

      this.Document.ClearParagraphsCache();
      return list;
    }

    #endregion

    #region Private Methods

    private void ClearMainParentContainerCache( Paragraph p )
    {
      if( p == null )
        return;

      var mainParentContainer = p.GetMainParentContainer();
      if( mainParentContainer != null )
      {
        mainParentContainer.ClearParagraphsCache();
      }
    }

    private void AddParagraphInCache( Paragraph p )
    {
      if( p == null )
        return;

      var mainParentContainer = p.GetMainParentContainer();
      if( mainParentContainer != null )
      {
        mainParentContainer.AddParagraphInCache( p );
        mainParentContainer.NeedRefreshParagraphIndexes = true;
      }
    }

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

    private void AddPicturesInPackage( Table t )
    {
      if( t == null )
        return;

      // Convert the path of this mainPart to its equilivant rels file path.
      var path = this.PackagePart.Uri.OriginalString.Replace( "/word/", "" );
      var rels_path = new Uri( "/word/_rels/" + path + ".rels", UriKind.Relative );

      // Check to see if the rels file exists and create it if not.
      if( !Document._package.PartExists( rels_path ) )
      {
        HelperFunctions.CreateRelsPackagePart( this.Document, rels_path );
      }

      foreach( var p in t.Pictures )
      {
        // Check to see if a rel for this Picture exists, create it if not.
        var rel_Id = HelperFunctions.GetOrGenerateRel( p._img._pr.TargetUri, this.PackagePart, TargetMode.Internal, Document.RelationshipImage );

        // Extract the attribute id from the Pictures Xml.
        var embed_id =
        (
            from e in p.Xml.Elements().Last().Descendants()
            where e.Name.LocalName.Equals( "blip" )
            select e.Attribute( XName.Get( "embed", Document.r.NamespaceName ) )
        ).Single();

        // Set its value to the Pictures relationships id.
        embed_id.SetValue( rel_Id );

        // Extract the attribute id from the Pictures Xml.
        var docPr =
        (
            from e in p.Xml.Elements().Last().Descendants()
            where e.Name.LocalName.Equals( "docPr" )
            select e
        ).Single();

        // Set its value to a unique id.
        docPr.SetAttributeValue( "id", this.Document.GetNextFreeDocPrId().ToString() );
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
