/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2023 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Linq;
using System.ComponentModel;

namespace Xceed.Document.NET
{
  public class Section : Container
  {
    #region Namespaces

    static internal XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static internal XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    #endregion

    #region Private Members

    private static float _pageSizeMultiplier = 20.0f;
    private NoteProperties _footnoteProperties;
    private NoteProperties _endnoteProperties;
    private PageNumberType _pageNumberType;

#endregion

    #region Properties

    public bool DifferentFirstPage
    {
      get
      {
        var titlePg = this.Xml.Element( w + "titlePg" );
        return ( titlePg != null );
      }

      set
      {
        var titlePg = this.Xml.Element( w + "titlePg" );
        if( titlePg == null )
        {
          if( value )
          {
            this.Xml.Add( new XElement( w + "titlePg", string.Empty ) );
          }
        }
        else
        {
          if( !value )
          {
            titlePg.Remove();
          }
        }
      }
    }

    public Footers Footers
    {
      get;
      internal set;
    }

    public NoteProperties FootnoteProperties
    {
      get
      {
        if( _footnoteProperties != null )
          return _footnoteProperties;

        _footnoteProperties = this.GetNoteProperties( "footnotePr" );
        if( _footnoteProperties != null )
        {
          _footnoteProperties.PropertyChanged += this.FootnoteProperties_PropertyChanged;
        }

        return _footnoteProperties;
      }

      set
      {
        if( _footnoteProperties != null )
        {
          _footnoteProperties.PropertyChanged -= this.FootnoteProperties_PropertyChanged;
        }

        _footnoteProperties = value;

        if( _footnoteProperties != null )
        {
          _footnoteProperties.PropertyChanged += this.FootnoteProperties_PropertyChanged;
        }

        this.SetNoteProperties( "footnotePr", _footnoteProperties );
      }
    }

    public NoteProperties EndnoteProperties
    {
      get
      {
        if( _endnoteProperties != null )
          return _endnoteProperties;

        _endnoteProperties = this.GetNoteProperties( "endnotePr" );
        if( _endnoteProperties != null )
        {
          _endnoteProperties.PropertyChanged += this.EndnoteProperties_PropertyChanged;
        }

        return _endnoteProperties;
      }

      set
      {
        if( _endnoteProperties != null )
        {
          _endnoteProperties.PropertyChanged -= this.EndnoteProperties_PropertyChanged;
        }

        _endnoteProperties = value;

        if( _endnoteProperties != null )
        {
          _endnoteProperties.PropertyChanged += this.EndnoteProperties_PropertyChanged;
        }

        this.SetNoteProperties( "endnotePr", _endnoteProperties );
      }
    }

    public Headers Headers
    {
      get;
      internal set;
    }

    /// <summary>
    /// Bottom margin in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginBottom
    {
      get
      {
        return this.GetMarginAttribute( XName.Get( "bottom", w.NamespaceName ) );
      }

      set
      {
        this.SetMarginAttribute( XName.Get( "bottom", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Footer margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginFooter
    {
      get
      {
        return this.GetMarginAttribute( XName.Get( "footer", w.NamespaceName ) );
      }
      set
      {
        this.SetMarginAttribute( XName.Get( "footer", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Header margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginHeader
    {
      get
      {
        return this.GetMarginAttribute( XName.Get( "header", w.NamespaceName ) );
      }
      set
      {
        this.SetMarginAttribute( XName.Get( "header", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Left margin in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginLeft
    {
      get
      {
        return this.GetMarginAttribute( XName.Get( "left", w.NamespaceName ) );
      }

      set
      {
        this.SetMarginAttribute( XName.Get( "left", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Right margin in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginRight
    {
      get
      {
        return this.GetMarginAttribute( XName.Get( "right", w.NamespaceName ) );
      }

      set
      {
        this.SetMarginAttribute( XName.Get( "right", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Top margin in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginTop
    {
      get
      {
        return this.GetMarginAttribute( XName.Get( "top", w.NamespaceName ) );
      }

      set
      {
        this.SetMarginAttribute( XName.Get( "top", w.NamespaceName ), value );
      }
    }

    public bool MirrorMargins
    {
      get
      {
        return this.GetMirrorMargins( XName.Get( "mirrorMargins", w.NamespaceName ) );
      }
      set
      {
        this.SetMirrorMargins( XName.Get( "mirrorMargins", w.NamespaceName ), value );
      }
    }

    public Borders PageBorders
    {
      get
      {
        var pgBorders = this.Xml.Element( XName.Get( "pgBorders", w.NamespaceName ) );
        if( pgBorders != null )
        {
          var pageBorders = new Borders();
          var top = pgBorders.Element( XName.Get( "top", w.NamespaceName ) );
          if( top != null )
          {
            pageBorders.Top = HelperFunctions.GetBorderFromXml( top );
          }
          var bottom = pgBorders.Element( XName.Get( "bottom", w.NamespaceName ) );
          if( bottom != null )
          {
            pageBorders.Bottom = HelperFunctions.GetBorderFromXml( bottom );
          }
          var left = pgBorders.Element( XName.Get( "left", w.NamespaceName ) );
          if( left != null )
          {
            pageBorders.Left = HelperFunctions.GetBorderFromXml( left );
          }
          var right = pgBorders.Element( XName.Get( "right", w.NamespaceName ) );
          if( right != null )
          {
            pageBorders.Right = HelperFunctions.GetBorderFromXml( right );
          }

          return pageBorders;
        }

        return null;
      }

      set
      {
        var pgBorders = this.Xml.Element( XName.Get( "pgBorders", w.NamespaceName ) );
        if( pgBorders == null )
        {
          pgBorders = new XElement( XName.Get( "pgBorders", Document.w.NamespaceName ) );
          this.Xml.Add( pgBorders );
        }

        if( value != null )
        {
          var topBorderValue = this.GetBorderAttributes( value.Top );
          if( topBorderValue != null )
          {
            pgBorders.Add( new XElement( XName.Get( "top", Document.w.NamespaceName ), topBorderValue ) );
          }
          var bottomBorderValue = this.GetBorderAttributes( value.Bottom );
          if( bottomBorderValue != null )
          {
            pgBorders.Add( new XElement( XName.Get( "bottom", Document.w.NamespaceName ), bottomBorderValue ) );
          }
          var leftBorderValue = this.GetBorderAttributes( value.Left );
          if( leftBorderValue != null )
          {
            pgBorders.Add( new XElement( XName.Get( "left", Document.w.NamespaceName ), leftBorderValue ) );
          }
          var rightBorderValue = this.GetBorderAttributes( value.Right );
          if( rightBorderValue != null )
          {
            pgBorders.Add( new XElement( XName.Get( "right", Document.w.NamespaceName ), rightBorderValue ) );
          }
        }
      }
    }

    /// <summary>
    /// Page height in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float PageHeight
    {
      get
      {
        var pgSz = this.Xml.Element( XName.Get( "pgSz", w.NamespaceName ) );
        if( pgSz != null )
        {
          var w = pgSz.Attribute( XName.Get( "h", Document.w.NamespaceName ) );
          if( w != null )
          {
            float f;
            if( HelperFunctions.TryParseFloat( w.Value, out f ) )
              return (int)( f / _pageSizeMultiplier );
          }
        }

        return ( 15840.0f / _pageSizeMultiplier );
      }

      set
      {
        var pgSz = this.Xml.Element( XName.Get( "pgSz", w.NamespaceName ) );
        if( pgSz != null )
        {
          pgSz.SetAttributeValue( XName.Get( "h", w.NamespaceName ), value * Convert.ToInt32( _pageSizeMultiplier ) );
        }
      }
    }

    public PageLayout PageLayout
    {
      get;
      private set;
    }

    public PageNumberType PageNumberType
    {
      get
      {
        var pgNumType = this.Xml.Element(XName.Get("pgNumType", w.NamespaceName));
        _pageNumberType = new PageNumberType();

        if (pgNumType != null)
        {
          var start = pgNumType.Attribute(XName.Get("start", Document.w.NamespaceName));
          var chapStyle = pgNumType.Attribute(XName.Get("chapStyle", Document.w.NamespaceName));
          var fmt = pgNumType.Attribute(XName.Get("fmt", Document.w.NamespaceName));
          var chapSep = pgNumType.Attribute(XName.Get("chapSep", Document.w.NamespaceName));


          if (start != null)
          {
            int i;
            if (HelperFunctions.TryParseInt(start.Value, out i))
              _pageNumberType.PageNumberStart = i;
          }

          if (chapStyle != null)
          {
            int i;
            if (HelperFunctions.TryParseInt(chapStyle.Value, out i))
              _pageNumberType.ChapterStyle = i;
          }

          if (fmt != null)
          {
            if (fmt.Value.Equals("decimal"))
            {
              _pageNumberType.PageNumberFormat = NumberingFormat.decimalNormal;
            }
            else
            {
              _pageNumberType.PageNumberFormat = Enum.GetValues(typeof(NumberingFormat)).Cast<NumberingFormat>().FirstOrDefault(x => x.ToString().Equals(fmt.Value));
            }
          }

          if (chapSep != null)
          {
            _pageNumberType.ChapterNumberSeperator = Enum.GetValues(typeof(ChapterSeperator)).Cast<ChapterSeperator>().FirstOrDefault(x => x.ToString().Equals(chapSep.Value)) ;
          }
        }

        _pageNumberType.PropertyChanged += this.PageNumberType_PropertyChanged;

        return _pageNumberType;
      }

      set
      {
        if (_pageNumberType != null)
        {
          _pageNumberType.PropertyChanged -= this.PageNumberType_PropertyChanged;
        }

        _pageNumberType = value;

        if (_pageNumberType != null)
        {
          _pageNumberType.PropertyChanged += this.PageNumberType_PropertyChanged;
        }

        this.UpdatePageNumberTypeXml();
      }
    }

    /// <summary>
    /// Page width in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float PageWidth
    {
      get
      {
        var pgSz = this.Xml.Element( XName.Get( "pgSz", w.NamespaceName ) );
        if( pgSz != null )
        {
          var w = pgSz.Attribute( XName.Get( "w", Document.w.NamespaceName ) );
          if( w != null )
          {
            float f;
            if( HelperFunctions.TryParseFloat( w.Value, out f ) )
              return (int)( f / _pageSizeMultiplier );
          }
        }

        return ( 12240.0f / _pageSizeMultiplier );
      }

      set
      {
        var pgSz = this.Xml.Element( XName.Get( "pgSz", w.NamespaceName ) );
        pgSz?.SetAttributeValue( XName.Get( "w", w.NamespaceName ), value * Convert.ToInt32( _pageSizeMultiplier ) );
      }
    }

    public SectionBreakType SectionBreakType
    {
      get
      {
        var type = this.Xml.Element( XName.Get( "type", Document.w.NamespaceName ) );
        if( type != null )
        {
          var val = type.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );
          switch( val )
          {
            case "continuous":
              return SectionBreakType.continuous;
            case "evenPage":
              return SectionBreakType.evenPage;
            case "oddPage":
              return SectionBreakType.oddPage;              
          }
        }
        return SectionBreakType.defaultNextPage;
      }

      set
      {
        var breakType = "nextPage";
        switch( value )
        {
          case SectionBreakType.continuous:
            breakType = "continuous";
            break;
          case SectionBreakType.evenPage:
            breakType = "evenPage";
            break;
          case SectionBreakType.oddPage:
            breakType = "oddPage";
            break;
        }

        var type = this.Xml.Element( XName.Get( "type", Document.w.NamespaceName ) );
        if( type == null )
        {
          this.Xml.Add( new XElement( XName.Get( "type", Document.w.NamespaceName ) ) );
          type = this.Xml.Element( XName.Get( "type", Document.w.NamespaceName ) );
        }
        type.SetAttributeValue( XName.Get( "val", w.NamespaceName ), breakType );
      }
    }

    public List<Paragraph> SectionParagraphs
    {
      get; set;
    }

    public override List<Table> Tables
    {
      get
      {
        var tables = new List<Table>();

        var sectionParagraphs = this.SectionParagraphs;
        if( sectionParagraphs.Count > 0 )
        {
          var parentElement = sectionParagraphs[ 0 ].Xml.Parent;
          if( parentElement.Name.LocalName == "tc" )
          {
            while( ( parentElement != null ) && ( parentElement.Name.LocalName != "tbl" ) )
            {
              parentElement = parentElement.Parent;
            }

            if( parentElement != null )
            {
              var table = new Table( this.Document, parentElement, this.PackagePart );
              tables.Add( this.Document.Tables.FirstOrDefault( t => t.Index == table.Index ) );
            }
          }
        }

        foreach( var paragraph in sectionParagraphs )
        {
          if( paragraph.FollowingTables != null )
          {
            tables.AddRange( paragraph.FollowingTables );
          }
        }

        return tables;
      }
    }

    #endregion

    #region Constructors

    internal Section( Document document, XElement xml, IEnumerable<XElement> lastSectionsXml ) : base( document, xml )
    {
      this.PageLayout = new PageLayout( document, xml );

      var xmlCopy = new XElement( xml );

      if( lastSectionsXml != null )
      {
        var lastSectionsXmlList = lastSectionsXml.ToList();
        for( int i = lastSectionsXmlList.Count - 1; i >= 0; i-- )
        {
          var lastSectionElements = lastSectionsXmlList[i].Elements();
          foreach( var lastSectionElement in lastSectionElements )
          {
            if( ( xmlCopy.Element( lastSectionElement.Name ) == null )
                  && ( lastSectionElement.Name.LocalName != "headerReference" )
                  && ( lastSectionElement.Name.LocalName != "footerReference" ) )
            {
              xmlCopy.Add( lastSectionElement );
            }
          }
        }
      }

      // Add last section header/footer references to this section xml copy.
      this.UpdateXmlReferenceFromLastSection( xmlCopy, lastSectionsXml, true );
      this.UpdateXmlReferenceFromLastSection( xmlCopy, lastSectionsXml, false );

      // Create the Header/Footer container based on the xml copy.
      this.AddHeadersContainer( xmlCopy );
      this.AddFootersContainer( xmlCopy );

      this.PackagePart = document.PackagePart;
    }


    #endregion

    #region Overrides

    public override Section InsertSection( bool trackChanges )
    {
      return this.InsertSection( trackChanges, false );
    }

    public override Section InsertSectionPageBreak( bool trackChanges = false )
    {
      return this.InsertSection( trackChanges, true );
    }

    protected internal override void AddElementInXml( object element )
    {
      if( this.SectionParagraphs.Count() > 0 )
      {
        this.SectionParagraphs.Last().Xml.AddBeforeSelf( element );
      }
      else
      {
        this.Xml.AddBeforeSelf( element );
      }
    }

    #endregion

    #region Public Methods

    public void AddHeaders()
    {
      this.AddHeadersOrFootersXml( true );
    }

    public void AddFooters()
    {
      this.AddHeadersOrFootersXml( false );
    }

    public void Remove()
    {
      if( this.Document.Sections.Count == 1 )
        throw new InvalidOperationException( "Can't remove the last section of a document." );

      // Remove all possible table (in paragraph or starting the section).
      this.Tables.ForEach( t => t.RemoveInternal() );

      // Remove all section paragraphs.
      foreach( var paragraphToFind in this.SectionParagraphs )
      {
        var paragraph = this.Document.Paragraphs.FirstOrDefault( p => p.Equals( paragraphToFind ) );
        if( paragraph != null )
        {
          paragraph.Remove( false );
        }
      }

      this.DeleteHeadersOrFooters( true );
      this.DeleteHeadersOrFooters( false );

      this.Xml.Remove();

      // When removing last section of a document, move the preceding section at the end of body.
      if( this.Document.Sections.Last() == this )
      {
        var lastSectionXml = this.Document.Xml.Descendants( XName.Get( "sectPr", w.NamespaceName ) ).Last();
        this.Document.Xml.Add( lastSectionXml );
        lastSectionXml.Remove();
      }

      this.Document.UpdateCacheSections();
    }

    #endregion

    #region Internal Methods

    /// <summary>
    /// Adds a Header to a section.
    /// If the section already contains a Header it will be replaced.
    /// </summary>
    internal void AddHeadersOrFootersXml( bool b )
    {
      var element = b ? "hdr" : "ftr";
      var reference = b ? "header" : "footer";

      this.DeleteHeadersOrFooters( b );

      var sectPr = this.Xml;

      // Get all header Relationships in this document.
      var biggestHeader = this.GetBiggestHeaderFooter( reference );

      for( var i = biggestHeader + 1; i < biggestHeader + 4; i++ )
      {
        this.CreateHeaderFooterPackage( element, reference, i, sectPr, i % 3 );
      }

      if( b )
      {
        this.AddHeadersContainer( sectPr );
      }
      else
      {
        this.AddFootersContainer( sectPr );
      }
    }

    internal void AddHeader( int headerType )
    {
      var sectPr = this.Xml;

      var biggestHeader = this.GetBiggestHeaderFooter( "header" );

      this.CreateHeaderFooterPackage( "hdr", "header", biggestHeader + 1, sectPr, headerType );

      this.AddHeaderContainer( sectPr, headerType );
    }

    #endregion

    #region Private Methods

    private void UpdateXmlReferenceFromLastSection( XElement xml, IEnumerable<XElement> lastSectionsXml, bool isHeader )
    {
      if( ( xml == null ) || ( lastSectionsXml == null ) || ( lastSectionsXml.Count() == 0 ) )
        return;

      var references = xml.Elements( XName.Get( isHeader ? "headerReference" : "footerReference", w.NamespaceName ) );

      // First, Even, Odd(default)
      var definedReferenceTypes = new List<XElement>() { null, null, null };
      foreach( var r in references )
      {
        var rType = r.Attribute( w + "type" ).Value;
        switch( rType )
        {
          case "first":
            definedReferenceTypes[ 0 ] = r;
            break;
          case "even":
            definedReferenceTypes[ 1 ] = r;
            break;
          default:
            definedReferenceTypes[ 2 ] = r;
            break;
        }
      }

      // Current section do not have a reference, copy the one from preceding sections, if available.
      if( definedReferenceTypes.Any( r => r == null ) )
      {
        var lastSectionsXmlList = lastSectionsXml.ToList();
        for( int i = lastSectionsXmlList.Count - 1; i >= 0; i-- )
        {
          var lastSectionXml = lastSectionsXmlList[ i ];
          var lastSectionReferences = lastSectionXml.Elements( XName.Get( isHeader ? "headerReference" : "footerReference", w.NamespaceName ) );

          if( definedReferenceTypes[ 0 ] == null )
          {
            var lastSectionFirst = lastSectionReferences.FirstOrDefault( x => x.Attribute( w + "type" ).Value == "first" );
            if( lastSectionFirst != null )
            {
              xml.Add( lastSectionFirst );
              definedReferenceTypes[ 0 ] = lastSectionFirst;
            }
          }
          if( definedReferenceTypes[ 1 ] == null )
          {
            var lastSectionEven = lastSectionReferences.FirstOrDefault( x => x.Attribute( w + "type" ).Value == "even" );
            if( lastSectionEven != null )
            {
              xml.Add( lastSectionEven );
              definedReferenceTypes[ 1 ] = lastSectionEven;
            }
          }
          if( definedReferenceTypes[ 2 ] == null )
          {
            var lastSectionDefault = lastSectionReferences.FirstOrDefault( x => x.Attribute( w + "type" ).Value == "default" );
            if( lastSectionDefault != null )
            {
              xml.Add( lastSectionDefault );
              definedReferenceTypes[ 2 ] = lastSectionDefault;
            }
          }

          if( definedReferenceTypes.All( r => r != null ) )
            break;
        }
      }
    }

    private void AddHeadersContainer( XElement xml )
    {
      Debug.Assert( xml != null, "xml shouldn't be null." );

      this.Headers = new Headers();
      var headerReferences = xml.Elements( XName.Get( "headerReference", w.NamespaceName ) );

      foreach( var h in headerReferences )
      {
        var hId = h.Attribute( r + "id" ).Value;
        var hType = h.Attribute( w + "type" ).Value;

        // Get the Xml file for this Header.
        var partUri = this.Document.PackagePart.GetRelationship( hId ).TargetUri;

        // Weird problem with PackaePart API.
        if( !partUri.OriginalString.StartsWith( "/word/" ) )
        {
          partUri = new Uri( "/word/" + partUri.OriginalString, UriKind.Relative );
        }

        // Get the Part and open a stream to get the Xml file.
        var part = this.Document._package.GetPart( partUri );

        using( TextReader tr = new StreamReader( part.GetStream() ) )
        {
          var doc = XDocument.Load( tr );
          // Header extend Container.
          var header = new Header( this.Document, doc.Element( w + "hdr" ), part, hId );
          switch( hType )
          {
            case "even":
              this.Headers.Even = header;
              break;
            case "first":
              this.Headers.First = header;
              break;
            default:
              this.Headers.Odd = header;
              break;
          }
        }
      }
    }

    private void AddFootersContainer( XElement xml )
    {
      Debug.Assert( xml != null, "xml shouldn't be null." );

      this.Footers = new Footers();
      var footerReferences = xml.Elements( XName.Get( "footerReference", w.NamespaceName ) );

      foreach( var f in footerReferences )
      {
        var fId = f.Attribute( r + "id" ).Value;
        var fType = f.Attribute( w + "type" ).Value;

        // Get the Xml file for this Footer.
        var partUri = this.Document.PackagePart.GetRelationship( fId ).TargetUri;

        // Weird problem with PackaePart API.
        if( !partUri.OriginalString.StartsWith( "/word/" ) )
        {
          partUri = new Uri( "/word/" + partUri.OriginalString, UriKind.Relative );
        }

        // Get the Part and open a stream to get the Xml file.
        var part = this.Document._package.GetPart( partUri );

        using( TextReader tr = new StreamReader( part.GetStream() ) )
        {
          var doc = XDocument.Load( tr );
          // Footer extend Container.
          var footer = new Footer( this.Document, doc.Element( w + "ftr" ), part, fId );
          switch( fType )
          {
            case "even":
              this.Footers.Even = footer;
              break;
            case "first":
              this.Footers.First = footer;
              break;
            default:
              this.Footers.Odd = footer;
              break;
          }
        }
      }
    }

    private void DeleteHeadersOrFooters( bool b )
    {
      var reference = b ? "header" : "footer";

      // Remove headerReferences and footerReferences from Xml.
      var sectPr = this.Xml;
      var existingReferences = sectPr.Elements( XName.Get( string.Format( "{0}Reference", reference ), w.NamespaceName ) );
      existingReferences.Remove();

      // Get all header(or footer) Relationships in this document.
      var header_relationships = this.Document.PackagePart.GetRelationshipsByType( string.Format( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference ) );

      foreach( var header_relationship in header_relationships )
      {
        // Get the TargetUri for this Part.
        var header_uri = header_relationship.TargetUri;

        // Check to see if the document actually contains the Part.
        if( !header_uri.OriginalString.StartsWith( "/word/" ) )
        {
          header_uri = new Uri( "/word/" + header_uri.OriginalString, UriKind.Relative );
        }

        if( this.Document._package.PartExists( header_uri ) )
        {
          // Get all references to this Relationship in the document.
          var query =
          (
              from e in this.Document._mainDoc.Descendants( XName.Get( "body", w.NamespaceName ) ).Descendants()
              where ( e.Name.LocalName == string.Format( "{0}Reference", reference ) ) && ( e.Attribute( r + "id" ).Value == header_relationship.Id )
              select e
          );

          // Delete the part and relationship not used anymore.
          if( query.Count() == 0 )
          {
            // Delete the Part
            this.Document._package.DeletePart( header_uri );
            // Delete the Relationship.
            this.Document._package.DeleteRelationship( header_relationship.Id );
          }
        }
      }
    }

    private void AddHeaderContainer( XElement xml, int headerType )
    {
      Debug.Assert( xml != null, "xml shouldn't be null." );

      var headerReferences = xml.Elements( XName.Get( "headerReference", w.NamespaceName ) );

      foreach( var h in headerReferences )
      {
        var hId = h.Attribute( r + "id" ).Value;
        var hType = h.Attribute( w + "type" ).Value;

        if( ( ( hType == "even" ) && ( headerType == 1 ) )
          || ( ( hType == "first" ) && ( headerType == 2 ) )
          || ( ( hType != "first" ) && ( hType != "even" ) && ( headerType == 0 ) ) )
        {
          // Get the Xml file for this Header.
          var partUri = this.Document.PackagePart.GetRelationship( hId ).TargetUri;

          // Weird problem with PackaePart API.
          if( !partUri.OriginalString.StartsWith( "/word/" ) )
          {
            partUri = new Uri( "/word/" + partUri.OriginalString, UriKind.Relative );
          }

          // Get the Part and open a stream to get the Xml file.
          var part = this.Document._package.GetPart( partUri );

          using( var tr = new StreamReader( part.GetStream() ) )
          {
            var doc = XDocument.Load( tr );
            // Header extend Container.
            var header = new Header( this.Document, doc.Element( w + "hdr" ), part, hId );
            switch( hType )
            {
              case "even":
                this.Headers.Even = header;
                break;
              case "first":
                this.Headers.First = header;
                break;
              default:
                this.Headers.Odd = header;
                break;
            }
          }
          break;
        }
      }
    }

    private int GetBiggestHeaderFooter( string reference )
    {
      var biggestHeader = 0;
      var header_relationships = this.Document.PackagePart.GetRelationshipsByType( string.Format( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference ) );
      // Get biggest headerX.xml.
      foreach( var header_relationship in header_relationships )
      {
        var header_uri = header_relationship.TargetUri;
        if( !header_uri.OriginalString.StartsWith( "/word/" ) )
        {
          header_uri = new Uri( "/word/" + header_uri.OriginalString, UriKind.Relative );
        }

        if( this.Document._package.PartExists( header_uri ) )
        {
          var resultString = Regex.Match( header_uri.OriginalString, @"\d+" ).Value;
          biggestHeader = Math.Max( biggestHeader, Int32.Parse( resultString ) );
        }
      }

      return biggestHeader;
    }

    private void CreateHeaderFooterPackage( string element, string reference, int index, XElement sectPr, int headerType )
    {
      var header_uri = string.Format( "/word/{0}{1}.xml", reference, index );

      var headerPart = this.Document._package.CreatePart( new Uri( header_uri, UriKind.Relative ), string.Format( "application/vnd.openxmlformats-officedocument.wordprocessingml.{0}+xml", reference ), CompressionOption.Normal );
      var headerRelationship = this.Document.PackagePart.CreateRelationship( headerPart.Uri, TargetMode.Internal, string.Format( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference ) );

      XDocument header;

      // Load the document part into a XDocument object
      using( TextReader tr = new StreamReader( headerPart.GetStream( FileMode.Create, FileAccess.ReadWrite ) ) )
      {
        header = XDocument.Parse
        ( string.Format( @"<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>
                       <w:{0} xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">
                         <w:p w:rsidR=""009D472B"" w:rsidRDefault=""009D472B"">
                           <w:pPr>
                             <w:pStyle w:val=""{1}"" />
                           </w:pPr>
                         </w:p>
                       </w:{0}>", element, reference )
        );
      }

      // Save the main document
      using( TextWriter tw = new StreamWriter( new PackagePartStream( headerPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        header.Save( tw, SaveOptions.None );
      }

      string type;
      switch( headerType )
      {
        case 0:
          type = "default";
          break;
        case 1:
          type = "even";
          break;
        case 2:
          type = "first";
          break;
        default:
          throw new ArgumentOutOfRangeException();
      }

      sectPr.Add
      (
          new XElement
          (
              w + string.Format( "{0}Reference", reference ),
              new XAttribute( w + "type", type ),
              new XAttribute( r + "id", headerRelationship.Id )
          )
      );
    }

    private float GetMarginAttribute( XName name )
    {
      var pgMar = this.Xml.Element( XName.Get( "pgMar", w.NamespaceName ) );
      var top = pgMar?.Attribute( name );
      if( top != null )
      {
        float f;
        if( HelperFunctions.TryParseFloat( top.Value, out f ) )
          return (int)( f / _pageSizeMultiplier );
      }

      return 0;
    }

    private void SetMarginAttribute( XName xName, float value )
    {
      var pgMar = this.Xml.Element( XName.Get( "pgMar", w.NamespaceName ) );
      var top = pgMar?.Attribute( xName );
      top?.SetValue( value * Convert.ToInt32( _pageSizeMultiplier ) );
    }

    private bool GetMirrorMargins( XName name )
    {
      var mirrorMargins = this.Xml.Element( XName.Get( "mirrorMargins", Document.w.NamespaceName ) );
      return ( mirrorMargins != null );
    }

    private void SetMirrorMargins( XName name, bool value )
    {
      var mirrorMargins = this.Xml.Element( XName.Get( "mirrorMargins", Document.w.NamespaceName ) );
      if( mirrorMargins == null )
      {
        this.Xml.Add( new XElement( w + "mirrorMargins", string.Empty ) );
      }
      else
      {
        if( !value )
        {
          mirrorMargins.Remove();
        }
      }
    }

    private object[] GetBorderAttributes( Border border )
    {
      if( border == null )
        return null;

      return new object[] { new XAttribute( XName.Get( "color", Document.w.NamespaceName ), border.Color.ToHex() ),
                            new XAttribute( XName.Get( "space", Document.w.NamespaceName ), border.Space ),
                            new XAttribute( XName.Get( "sz", Document.w.NamespaceName ),  Border.GetNumericSize( border.Size ) ),
                            new XAttribute( XName.Get( "val", Document.w.NamespaceName ), border.Tcbs.ToString().Remove(0, 5) )
                          };
    }

    private Section InsertSection( bool trackChanges, bool isPageBreak )
    {
      bool isLastSection = ( this.Document.Sections.Last() == this );

      // Save any modified header/footer so that the new section can access it.
      this.Document.SaveHeadersFooters();

      var sctPr = new XElement( this.Xml );
      if( !isPageBreak )
      {
        sctPr.Add( new XElement( XName.Get( "type", Document.w.NamespaceName ), new XAttribute( Document.w + "val", "continuous" ) ) );
      }

      if( isLastSection )
      {
        var currentSection = new XElement( XName.Get( "p", Document.w.NamespaceName ), new XElement( XName.Get( "pPr", Document.w.NamespaceName ), sctPr ) );
        if( this.SectionParagraphs.Count > 0 )
        {
          this.SectionParagraphs.Last().Xml.AddAfterSelf( currentSection );
        }
        else
        {
          this.Xml.AddBeforeSelf( currentSection );
        }

        this.Xml.Remove();
        this.Xml = currentSection;

        var newSection = sctPr;
        if( trackChanges )
        {
          newSection = HelperFunctions.CreateEdit( EditType.ins, DateTime.Now, newSection );
        }

        currentSection.AddAfterSelf( newSection );
      }
      else
      {
        var newSection = new XElement( XName.Get( "p", Document.w.NamespaceName ), new XElement( XName.Get( "pPr", Document.w.NamespaceName ), sctPr ) );
        if( trackChanges )
        {
          newSection = HelperFunctions.CreateEdit( EditType.ins, DateTime.Now, newSection );
        }

        this.SectionParagraphs.Last().Xml.AddAfterSelf( newSection );
      }      

      // Update the _cachedSection by reading the Xml to build new Sections.
      this.Document.UpdateCacheSections();

      return this.Document.Sections.LastOrDefault();
    }

    private NoteProperties GetNoteProperties( string propertiesType )
    {
      var notePr = this.Xml.Element( XName.Get( propertiesType, Document.w.NamespaceName ) );
      if( notePr != null )
      {
        var numberFormat = NoteNumberFormat.number;
        var numberStart = 1;

        var numFmt = notePr.Element( XName.Get( "numFmt", Document.w.NamespaceName ) );
        if( numFmt != null )
        {
          var val = numFmt.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );
          if( !string.IsNullOrEmpty( val ) )
          {
            NoteNumberFormat enumNumberFormat;
            numberFormat = Enum.TryParse( val, out enumNumberFormat ) ? enumNumberFormat : NoteNumberFormat.number;
          }
        }

        var numStart = notePr.Element( XName.Get( "numStart", Document.w.NamespaceName ) );
        if( numStart != null )
        {
          var val = numStart.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );
          if( !string.IsNullOrEmpty( val ) )
          {
            numberStart = int.Parse( val );
          }
        }

        return new NoteProperties()
        {
          NumberFormat = numberFormat,
          NumberStart = numberStart
        };
      }

      return null;
    }

    private void SetNoteProperties( string propertiesType, NoteProperties newValue )
    {
      if( newValue != null )
      {
        var notePr = this.Xml.Element( XName.Get( propertiesType, Document.w.NamespaceName ) );
        if( notePr == null )
        {
          this.Xml.Add( new XElement( XName.Get( propertiesType, Document.w.NamespaceName ) ) );
          notePr = this.Xml.Element( XName.Get( propertiesType, Document.w.NamespaceName ) );
        }

        var numFmt = notePr.Element( XName.Get( "numFmt", Document.w.NamespaceName ) );
        if( numFmt == null )
        {
          notePr.Add( new XElement( XName.Get( "numFmt", Document.w.NamespaceName ) ) );
          numFmt = notePr.Element( XName.Get( "numFmt", Document.w.NamespaceName ) );
        }

        numFmt.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), newValue.NumberFormat );

        var numStart = notePr.Element( XName.Get( "numStart", Document.w.NamespaceName ) );
        if( numStart == null )
        {
          notePr.Add( new XElement( XName.Get( "numStart", Document.w.NamespaceName ) ) );
          numStart = notePr.Element( XName.Get( "numStart", Document.w.NamespaceName ) );
        }

        numStart.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), newValue.NumberStart );
      }
      else
      {
        var notePr = this.Xml.Element( XName.Get( propertiesType, Document.w.NamespaceName ) );
        if( notePr != null )
        {
          notePr.Remove();
        }
      }
    }
    private void UpdatePageNumberTypeXml()
    {
      var pgNumType = this.Xml.Element(XName.Get("pgNumType", w.NamespaceName));

      if (pgNumType == null)
      {
        this.Xml.Add(new XElement(XName.Get("pgNumType", w.NamespaceName)));
        pgNumType = this.Xml.Element(XName.Get("pgNumType", w.NamespaceName));
      }

      if(_pageNumberType != null)
      {
        pgNumType.SetAttributeValue(XName.Get("chapStyle", w.NamespaceName), _pageNumberType.ChapterStyle);

        if (_pageNumberType.PageNumberStart.HasValue && _pageNumberType.PageNumberStart.Value > 0)
        {
          pgNumType.SetAttributeValue(XName.Get("start", w.NamespaceName), _pageNumberType.PageNumberStart);
        }
        else
        {
          pgNumType.SetAttributeValue(XName.Get("start", w.NamespaceName), null);
        }

        pgNumType.SetAttributeValue(XName.Get("chapSep", w.NamespaceName), _pageNumberType.ChapterNumberSeperator);

        if (_pageNumberType.PageNumberFormat == NumberingFormat.decimalNormal)
        {
          pgNumType.SetAttributeValue(XName.Get("fmt", w.NamespaceName), "decimal");
        }
        else
        {
          pgNumType.SetAttributeValue(XName.Get("fmt", w.NamespaceName), _pageNumberType.PageNumberFormat);
        }
      }
    }


    #endregion

    #region Event Handlers

    private void FootnoteProperties_PropertyChanged( object sender, PropertyChangedEventArgs e )
    {
      this.SetNoteProperties( "footnotePr", _footnoteProperties );
    }

    private void EndnoteProperties_PropertyChanged( object sender, PropertyChangedEventArgs e )
    {
      this.SetNoteProperties( "endnotePr", _endnoteProperties );
    }

    private void PageNumberType_PropertyChanged(object sender, PropertyChangedEventArgs e)
    {
      this.UpdatePageNumberTypeXml();
    }

    #endregion
  }
}
