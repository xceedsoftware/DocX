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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using System.Globalization;
using System.Diagnostics;
using System.IO;

namespace Xceed.Document.NET
{
  public class Paragraph : InsertBeforeOrAfter
  {
    #region Internal Members

    // The Append family of functions use this List to apply style.
    internal List<XElement> _runs;
    internal int _startIndex, _endIndex;
    internal List<XElement> _styles = new List<XElement>();

    internal const float DefaultSingleLineSpacing = 12f;
    internal static float DefaultLineSpacing = Paragraph.DefaultSingleLineSpacing;
    internal const string HyperlinkRelation = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    #endregion

    #region Private Members

    // This paragraphs text alignment
    private Alignment alignment;
    // A collection of field type DocProperty.
    private List<DocProperty> docProperties;
    private Direction direction;
    private float indentationFirstLine;
    private float indentationHanging;
    private float indentationBefore;
    private float indentationAfter = 0.0f;
    private List<Table> followingTables;
    private List<FormattedText> _magicText;
    private static int bookmarkIdCounter = 0;
    private static float DefaultLineSpacingAfter = 0f;
    private static float DefaultLineSpacingBefore = 0f;
    private static bool DefaultLineRuleAuto = false;

    private static float DefaultIndentationFirstLine = 0f;
    private static float DefaultIndentationHanging = 0f;
    private static float DefaultIndentationBefore = 0f;
    private static float DefaultIndentationAfter = 0f;

    private static float DefaultImageHorizontalResolution = 96f;
    private static float DefaultImageVerticalResolution = 96f;

    private bool m_removed = false; // this will be used when a paragraph is removed via RemoveText method

    #endregion

    #region Private Properties
    private XElement ParagraphNumberPropertiesBacker
    {
      get; set;
    }

    private int? IndentLevelBacker
    {
      get; set;
    }

    #endregion

    #region Internal Properties

    internal bool? IsListItemBacker
    {
      get; set;
    }

    internal int OutlineLevel
    {
      get
      {
        var pPr = GetOrCreate_pPr();
        var outlineLvl = pPr.Element( XName.Get( "outlineLvl", Document.w.NamespaceName ) );
        if( outlineLvl != null )
        {
          var val = outlineLvl.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
          return ( val != null ) ? Int32.Parse( val.Value ) : 0;
        }

        return 0;
      }
    }

    #endregion

    #region Public Properties










    public ContainerType ParentContainer
    {
      get; set;
    }
    public ListItemType ListItemType
    {
      get; set;
    }

    public List<Picture> Pictures
    {
      get
      {
        if( Xml == null )
        {
          return new List<Picture>();
        }

        var pictures = this.GetPictures( "drawing", "blip", "embed" );
        var shapes = this.GetPictures( "pict", "imagedata", "id" );

        foreach( Picture pict in shapes )
        {
          pictures.Add( pict );
        }

        return pictures;
      }
    }

    public List<Hyperlink> Hyperlinks
    {
      get
      {
        var hyperlinks = new List<Hyperlink>();

        var hyperlink_elements =
        (
            from h in Xml.Descendants()
            where ( h.Name.LocalName == "hyperlink" || h.Name.LocalName == "instrText" )
            select h
        ).ToList();

        foreach( XElement he in hyperlink_elements )
        {
          if( he.Name.LocalName == "hyperlink" )
          {
            try
            {
              var h = new Hyperlink( this.Document, this.PackagePart, he );
              h.PackagePart = this.PackagePart;
              hyperlinks.Add( h );
            }
            catch( Exception )
            {
            }
          }
          else
          {
            // Find the parent run, no matter how deeply nested we are.
            XElement e = he;
            while( e.Name.LocalName != "r" )
            {
              e = e.Parent;
            }

            // Take every element until we reach w:fldCharType="end"
            var hyperlink_runs = new List<XElement>();
            foreach( XElement r in e.ElementsAfterSelf( XName.Get( "r", Document.w.NamespaceName ) ) )
            {
              // Add this run to the list.
              hyperlink_runs.Add( r );

              var fldChar = r.Descendants( XName.Get( "fldChar", Document.w.NamespaceName ) ).SingleOrDefault<XElement>();
              if( fldChar != null )
              {
                var fldCharType = fldChar.Attribute( XName.Get( "fldCharType", Document.w.NamespaceName ) );
                if( fldCharType != null && fldCharType.Value.Equals( "end", StringComparison.CurrentCultureIgnoreCase ) )
                {
                  try
                  {
                    var h = new Hyperlink( Document, he, hyperlink_runs );
                    h.PackagePart = this.PackagePart;
                    hyperlinks.Add( h );
                  }
                  catch( Exception )
                  {
                  }

                  break;
                }
              }
            }
          }
        }

        return hyperlinks;
      }
    }

    [Obsolete( "This property is obsolete and should no longer be used. Use StyleId instead." )]
    public string StyleName
    {
      get
      {
        return this.StyleId;
      }
      set
      {
        this.StyleId = value;
      }
    }

    public string StyleId
    {
      get
      {
        var element = this.GetOrCreate_pPr();
        var styleElement = element.Element( XName.Get( "pStyle", Document.w.NamespaceName ) );
        var attr = styleElement?.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
        if( !string.IsNullOrEmpty( attr?.Value ) )
        {
          return attr.Value;
        }

        return this.Document.GetNormalStyleId();
      }
      set
      {
        if( string.IsNullOrEmpty( value ) )
        {
          value = this.Document.GetNormalStyleId();
        }
        var element = this.GetOrCreate_pPr();
        var styleElement = element.Element( XName.Get( "pStyle", Document.w.NamespaceName ) );
        if( styleElement == null )
        {
          element.Add( new XElement( XName.Get( "pStyle", Document.w.NamespaceName ) ) );
          styleElement = element.Element( XName.Get( "pStyle", Document.w.NamespaceName ) );
        }
        styleElement.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), value );

        this.AddParagraphStyleIfNotPresent( this.StyleId );
      }
    }

    public List<DocProperty> DocumentProperties
    {
      get
      {
        return docProperties;
      }
    }

    public Direction Direction
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();
        XElement bidi = pPr.Element( XName.Get( "bidi", Document.w.NamespaceName ) );
        return bidi == null ? Direction.LeftToRight : Direction.RightToLeft;
      }

      set
      {
        direction = value;

        XElement pPr = GetOrCreate_pPr();
        XElement bidi = pPr.Element( XName.Get( "bidi", Document.w.NamespaceName ) );

        if( direction == Direction.RightToLeft )
        {
          if( bidi == null )
            pPr.Add( new XElement( XName.Get( "bidi", Document.w.NamespaceName ) ) );
        }

        else
        {
          bidi?.Remove();
        }
      }
    }

    public float IndentationFirstLine
    {
      get
      {
        GetOrCreate_pPr();
        XElement ind = GetOrCreate_pPr_ind();
        XAttribute firstLine = ind.Attribute( XName.Get( "firstLine", Document.w.NamespaceName ) );

        if( firstLine != null )
          return float.Parse( firstLine.Value ) / 20f;

        return Paragraph.DefaultIndentationFirstLine;
      }

      set
      {
        if( IndentationFirstLine != value )
        {
          indentationFirstLine = value;

          GetOrCreate_pPr();
          XElement ind = GetOrCreate_pPr_ind();

          // Paragraph can either be firstLine or hanging (Remove hanging).
          XAttribute hanging = ind.Attribute( XName.Get( "hanging", Document.w.NamespaceName ) );
          hanging?.Remove();

          string indentation = ( indentationFirstLine * 20f ).ToString( CultureInfo.InvariantCulture );
          XAttribute firstLine = ind.Attribute( XName.Get( "firstLine", Document.w.NamespaceName ) );
          if( firstLine != null )
            firstLine.Value = indentation;
          else
            ind.Add( new XAttribute( XName.Get( "firstLine", Document.w.NamespaceName ), indentation ) );
        }
      }
    }

    public float IndentationHanging
    {
      get
      {
        GetOrCreate_pPr();
        var ind = GetOrCreate_pPr_ind();
        var hanging = ind.Attribute( XName.Get( "hanging", Document.w.NamespaceName ) );

        if( hanging != null )
          return float.Parse( hanging.Value ) / 20f;

        return Paragraph.DefaultIndentationHanging;
      }

      set
      {
        if( IndentationHanging != value )
        {
          indentationHanging = value;

          GetOrCreate_pPr();
          var ind = GetOrCreate_pPr_ind();

          // Paragraph can either be firstLine or hanging (Remove firstLine).
          var firstLine = ind.Attribute( XName.Get( "firstLine", Document.w.NamespaceName ) );
          if( firstLine != null )
          {
            firstLine.Remove();
          }

          var indentationValue = ( indentationHanging * 20 );
          var indentation = indentationValue.ToString( CultureInfo.InvariantCulture );
          var hanging = ind.Attribute( XName.Get( "hanging", Document.w.NamespaceName ) );
          if( hanging != null )
          {
            hanging.Value = indentation;
          }
          else
          {
            ind.Add( new XAttribute( XName.Get( "hanging", Document.w.NamespaceName ), indentation ) );
          }
          IndentationBefore = indentationHanging;
        }
      }
    }

    public float IndentationBefore
    {
      get
      {
        GetOrCreate_pPr();
        var ind = GetOrCreate_pPr_ind();

        var left = ind.Attribute( XName.Get( "left", Document.w.NamespaceName ) );
        if( left != null )
          return float.Parse( left.Value ) / 20f;

        return Paragraph.DefaultIndentationBefore;
      }

      set
      {
        if( IndentationBefore != value )
        {
          indentationBefore = value;

          GetOrCreate_pPr();
          var ind = GetOrCreate_pPr_ind();

          var indentation = ( indentationBefore * 20f ).ToString( CultureInfo.InvariantCulture );

          var left = ind.Attribute( XName.Get( "left", Document.w.NamespaceName ) );
          if( left != null )
          {
            left.Value = indentation;
          }
          else
          {
            ind.Add( new XAttribute( XName.Get( "left", Document.w.NamespaceName ), indentation ) );
          }
        }
      }
    }

    public float IndentationAfter
    {
      get
      {
        GetOrCreate_pPr();
        var ind = GetOrCreate_pPr_ind();

        var right = ind.Attribute( XName.Get( "right", Document.w.NamespaceName ) );
        if( right != null )
          return float.Parse( right.Value ) / 20f;

        return Paragraph.DefaultIndentationAfter;
      }

      set
      {
        if( IndentationAfter != value )
        {
          indentationAfter = value;

          GetOrCreate_pPr();
          var ind = GetOrCreate_pPr_ind();

          var indentation = ( indentationAfter * 20f ).ToString( CultureInfo.InvariantCulture );

          var right = ind.Attribute( XName.Get( "right", Document.w.NamespaceName ) );
          if( right != null )
          {
            right.Value = indentation;
          }
          else
          {
            ind.Add( new XAttribute( XName.Get( "right", Document.w.NamespaceName ), indentation ) );
          }
        }
      }
    }

    public Alignment Alignment
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();
        XElement jc = pPr.Element( XName.Get( "jc", Document.w.NamespaceName ) );

        if( jc != null )
        {
          XAttribute a = jc.Attribute( XName.Get( "val", Document.w.NamespaceName ) );

          switch( a.Value.ToLower() )
          {
            case "left":
              return Xceed.Document.NET.Alignment.left;
            case "right":
              return Xceed.Document.NET.Alignment.right;
            case "center":
              return Xceed.Document.NET.Alignment.center;
            case "both":
              return Xceed.Document.NET.Alignment.both;
          }
        }

        return Xceed.Document.NET.Alignment.left;
      }

      set
      {
        alignment = value;

        XElement pPr = GetOrCreate_pPr();
        XElement jc = pPr.Element( XName.Get( "jc", Document.w.NamespaceName ) );

        if( jc == null )
          pPr.Add( new XElement( XName.Get( "jc", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), alignment.ToString() ) ) );
        else
          jc.Attribute( XName.Get( "val", Document.w.NamespaceName ) ).Value = alignment.ToString();
      }
    }

    public string Text
    {
      // Returns the underlying XElement's Value property.
      get
      {
        try
        {
          return HelperFunctions.GetText( Xml );
        }
        catch( Exception )
        {
          return null;
        }
      }
    }

    public List<FormattedText> MagicText
    {
      // Returns the underlying XElement's Value property.
      get
      {
        if( _magicText == null )
        {
          _magicText = HelperFunctions.GetFormattedText( this.Document, Xml );
        }
        return _magicText;
      }
    }

    public Paragraph CurrentCulture()
    {
      ApplyTextFormattingProperty( XName.Get( "lang", Document.w.NamespaceName ),
          string.Empty,
          new XAttribute( XName.Get( "val", Document.w.NamespaceName ), CultureInfo.CurrentCulture.Name ) );
      return this;
    }

    public List<Table> FollowingTables
    {
      get
      {
        return followingTables;
      }
      internal set
      {
        followingTables = value;
      }
    }

    public float LineSpacing
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();
        XElement spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

        if( spacing != null )
        {
          XAttribute line = spacing.Attribute( XName.Get( "line", Document.w.NamespaceName ) );
          if( line != null )
          {
            float f;

            if( HelperFunctions.TryParseFloat( line.Value, out f ) )
              return f / 20.0f;
          }
        }

        var rPr = pPr.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
        if( rPr != null )
        {
          var size = rPr.Element( XName.Get( "sz", Document.w.NamespaceName ) );
          if( size != null )
            return Convert.ToSingle( size.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) ) / 2;
        }

        return Paragraph.DefaultLineSpacing;
      }

      set
      {
        SpacingLine( value );
      }
    }

    public float LineSpacingBefore
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();

        XElement spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

        if( spacing != null )
        {
          if( this.IsBeforeAutoSpacing() )
            return 0f;

          var line = spacing.Attribute( XName.Get( "before", Document.w.NamespaceName ) );
          if( line != null )
          {
            float f;

            if( HelperFunctions.TryParseFloat( line.Value, out f ) )
              return f / 20.0f;
          }
        }

        return Paragraph.DefaultLineSpacingBefore;
      }

      set
      {
        SpacingBefore( value );
      }
    }

    public float LineSpacingAfter
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();
        XElement spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

        if( spacing != null )
        {
          if( this.IsAfterAutoSpacing() )
            return 0f;

          var line = spacing.Attribute( XName.Get( "after", Document.w.NamespaceName ) );
          if( line != null )
          {
            float f;

            if( HelperFunctions.TryParseFloat( line.Value, out f ) )
              return f / 20.0f;
          }
        }

        return Paragraph.DefaultLineSpacingAfter;
      }

      set
      {
        SpacingAfter( value );
      }
    }

    internal bool IsAfterAutoSpacing()
    {
      XElement pPr = GetOrCreate_pPr();
      XElement spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

      if( spacing != null )
      {
        var afterAutoSpacing = spacing.Attribute( XName.Get( "afterAutospacing", Document.w.NamespaceName ) );
        if( ( afterAutoSpacing != null ) && ( afterAutoSpacing.Value == "1" ) )
          return true;
      }
      return false;
    }

    internal bool IsBeforeAutoSpacing()
    {
      XElement pPr = GetOrCreate_pPr();
      XElement spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

      if( spacing != null )
      {
        var beforeAutospacing = spacing.Attribute( XName.Get( "beforeAutospacing", Document.w.NamespaceName ) );
        if( ( beforeAutospacing != null ) && ( beforeAutospacing.Value == "1" ) )
          return true;
      }
      return false;
    }

    public XElement ParagraphNumberProperties
    {
      get
      {
        return ParagraphNumberPropertiesBacker ?? ( ParagraphNumberPropertiesBacker = GetParagraphNumberProperties() );
      }
    }

    public bool IsListItem
    {
      get
      {
        IsListItemBacker = IsListItemBacker ?? ( ParagraphNumberProperties != null );
        return (bool)IsListItemBacker;
      }
    }

    public int? IndentLevel
    {
      get
      {
        if( !IsListItem )
          return null;

        if( IndentLevelBacker != null )
          return IndentLevelBacker;

        var ilvl = ParagraphNumberProperties.Descendants().FirstOrDefault( el => el.Name.LocalName == "ilvl" );
        return IndentLevelBacker = ( ilvl != null ) ? int.Parse( ilvl.GetAttribute( Document.w + "val" ) ) : 0;
      }
    }

    public bool IsKeepWithNext
    {
      get
      {
        var pPr = this.GetOrCreate_pPr();
        var keepNext = pPr.Element( XName.Get( "keepNext", Document.w.NamespaceName ) );

        return ( keepNext != null );
      }
    }

    public int StartIndex
    {
      get
      {
        var mainContainer = this.GetMainParentContainer();
        if( ( mainContainer != null ) && mainContainer.NeedRefreshParagraphIndexes && !mainContainer.PreventUpdateParagraphIndexes )
        {
          mainContainer.RefreshParagraphIndexes();
        }

        if( this.IsInMainContainer() )
          return _startIndex;

        var documentParagraph = this.GetContainerParagraphs()?.FirstOrDefault( p => p.Xml == this.Xml );
        if( documentParagraph != null )
          return documentParagraph._startIndex;

        return 0;
      }
    }

    public int EndIndex
    {
      get
      {
        var mainContainer = this.GetMainParentContainer();
        if( ( mainContainer != null ) && mainContainer.NeedRefreshParagraphIndexes && !mainContainer.PreventUpdateParagraphIndexes )
        {
          mainContainer.RefreshParagraphIndexes();
        }

        if( this.IsInMainContainer() )
          return _endIndex;

        var documentParagraph = this.GetContainerParagraphs()?.FirstOrDefault( p => p.Xml == this.Xml );
        if( documentParagraph != null )
          return documentParagraph._endIndex;

        return 1;
      }
    }


    public List<Footnote> Footnotes
    {
      get
      {
        var footnotes = new List<Footnote>();

        if( this.Xml == null )
        {
          return footnotes;
        }


        // Get all footnote reference elements in the paragraph
        var footnoteReferences = this.Xml.Descendants( XName.Get( "footnoteReference", Document.w.NamespaceName ) ).ToList();

        if( footnoteReferences != null && footnoteReferences.Count > 0 )
        {
          foreach( var footnoteReference in footnoteReferences )
          {
            var footnoteId = footnoteReference.GetAttribute( XName.Get( "id", Document.w.NamespaceName ) );

            if( !string.IsNullOrEmpty( footnoteId ) )
            {
              var documentFootnotes = this.Document.GetFootnotes();

              if( documentFootnotes != null )
              {
                var footnoteXml = documentFootnotes.Descendants()
                                                   .FirstOrDefault( f => f.Attribute( XName.Get( "id", Document.w.NamespaceName ) ) != null
                                                                      && f.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value == footnoteId );

                if( footnoteXml != null )
                {
                  var footnote = new Footnote( this.Document, this, this.Document._footnotesPart, footnoteXml );
                  footnotes.Add( footnote );
                }
              }
            }
          }
        }

        return footnotes;
      }
    }

    public List<Endnote> Endnotes
    {
      get
      {
        var endnotes = new List<Endnote>();

        if( this.Xml == null )
        {
          return endnotes;
        }

        // Get all endnote reference elements in the paragraph
        var endnoteReferences = this.Xml.Descendants( XName.Get( "endnoteReference", Document.w.NamespaceName ) ).ToList();

        if( endnoteReferences != null && endnoteReferences.Count > 0 )
        {
          foreach( var endnoteReference in endnoteReferences )
          {
            var endnoteId = endnoteReference.GetAttribute( XName.Get( "id", Document.w.NamespaceName ) );

            if( !string.IsNullOrEmpty( endnoteId ) )
            {
              var documentEndnotes = this.Document.GetEndnotes();

              if( documentEndnotes != null )
              {
                var endnoteXml = documentEndnotes.Descendants()
                                                 .FirstOrDefault( f => f.Attribute( XName.Get( "id", Document.w.NamespaceName ) ) != null
                                                                    && f.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value == endnoteId );

                if( endnoteXml != null )
                {
                  var endnote = new Endnote( this.Document, this, this.Document._endnotesPart, endnoteXml );
                  endnotes.Add( endnote );
                }
              }
            }
          }
        }

        return endnotes;
      }
    }

    #endregion

    #region Constructors

    internal Paragraph( Document document, XElement xml, int startIndex = -1, ContainerType parentContainerType = ContainerType.None ) : base( document, xml )
    {
      //_startIndex = startIndex;

      //var alternateContentValue = xml.DescendantsAndSelf().FirstOrDefault( x => x.Name.Equals( XName.Get( "AlternateContent", Document.mc.NamespaceName ) ) );
      //if( alternateContentValue != null )
      //{
      //  StringBuilder sb = new StringBuilder();
      //  HelperFunctions.GetTextRecursive( xml, ref sb );
      //  var text = sb.ToString();

      //  _endIndex = startIndex + Math.Max( 1, text.Length );
      //}
      //else
      //{
      //  _endIndex = startIndex + GetElementTextLength( xml );
      //}


      ParentContainer = parentContainerType;

      RebuildDocProperties();

      //var stylesElements = xml.Descendants( XName.Get( "pStyle", Document.w.NamespaceName ) );

      //if( stylesElements.Count() > 0 )
      //{
      //  Uri style_package_uri = new Uri( "/word/styles.xml", UriKind.Relative );
      //  PackagePart styles_document = document.package.GetPart( style_package_uri );

      //  using( TextReader tr = new StreamReader( styles_document.GetStream() ) )
      //  {
      //    XDocument style_document = XDocument.Load( tr );
      //    XElement styles_element = style_document.Element( XName.Get( "styles", Document.w.NamespaceName ) );

      //    var styles_element_ids = stylesElements.Select( e => e.Attribute( XName.Get( "val", Document.w.NamespaceName ) ).Value );

      //    //foreach(string id in styles_element_ids)
      //    //{
      //    //    var style = 
      //    //    (
      //    //        from d in styles_element.Descendants()
      //    //        let styleId = d.Attribute(XName.Get("styleId", Document.w.NamespaceName))
      //    //        let type = d.Attribute(XName.Get("type", Document.w.NamespaceName))
      //    //        where type != null && type.Value == "paragraph" && styleId != null && styleId.Value == id
      //    //        select d
      //    //    ).First();

      //    //    styles.Add(style);
      //    //} 
      //  }
      //}

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();
    }

    internal Paragraph( XElement newParagraph ) : base( null, newParagraph )
    {
    }

    #endregion

    #region Public Methods

    public override Table InsertTableBeforeSelf( Table t )
    {
      t = base.InsertTableBeforeSelf( t );
      t.PackagePart = this.PackagePart;

      return t;
    }

    public override Table InsertTableBeforeSelf( int rowCount, int columnCount )
    {
      return base.InsertTableBeforeSelf( rowCount, columnCount );
    }

    public override Table InsertTableAfterSelf( Table t )
    {
      t = base.InsertTableAfterSelf( t );
      t.PackagePart = this.PackagePart;

      if( this.ParentContainer == ContainerType.Cell )
      {
        t.InsertParagraphAfterSelf( "" );
      }

      return t;
    }

    public override Table InsertTableAfterSelf( int rowCount, int columnCount )
    {
      var t = base.InsertTableAfterSelf( rowCount, columnCount );

      if( this.ParentContainer == ContainerType.Cell )
      {
        t.InsertParagraphAfterSelf( "" );
      }

      return t;
    }

    public Picture ReplacePicture( Picture toBeReplaced, Picture replaceWith )
    {
      var document = this.Document;
      Picture replacePicture = null;

      if( toBeReplaced == null )
        return null;

      if( replaceWith != null )
      {
        var newDocPrId = document.GetNextFreeDocPrId();

        var xml = XElement.Parse( toBeReplaced.Xml.ToString() );

        foreach( var element in xml.Descendants( XName.Get( "docPr", Document.wp.NamespaceName ) ) )
        {
          element.SetAttributeValue( XName.Get( "id" ), newDocPrId );
        }

        foreach( var element in xml.Descendants( XName.Get( "blip", Document.a.NamespaceName ) ) )
        {
          element.SetAttributeValue( XName.Get( "embed", Document.r.NamespaceName ), replaceWith.Id );
        }

        replacePicture = new Picture( this.Document, xml, new Image( document, this.PackagePart.GetRelationship( replaceWith.Id ) ) );
        this.AppendPicture( replacePicture );
      }
      toBeReplaced.Remove();

      return replacePicture;
    }

    public override Paragraph InsertParagraphBeforeSelf( Paragraph p )
    {
      this.ValidateInsert();
      this.ClearContainerParagraphsCache();

      var p2 = base.InsertParagraphBeforeSelf( p );
      p2.PackagePart = this.PackagePart;

      this.NeedRefreshIndexes();
      this.InsertFollowingTables( p, false );

      return p2;
    }

    public override Paragraph InsertParagraphBeforeSelf( string text )
    {
      this.ValidateInsert();

      // Inserting a paragraph at a specific index needs an update of Paragraph cache.
      this.ClearContainerParagraphsCache();
      var p = base.InsertParagraphBeforeSelf( text );
      p.PackagePart = this.PackagePart;

      this.NeedRefreshIndexes();

      return p;
    }

    public override Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges )
    {
      this.ValidateInsert();
      this.ClearContainerParagraphsCache();

      var p = base.InsertParagraphBeforeSelf( text, trackChanges );
      p.PackagePart = this.PackagePart;

      this.NeedRefreshIndexes();

      return p;
    }

    public override Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges, Formatting formatting )
    {
      this.ValidateInsert();
      this.ClearContainerParagraphsCache();

      var p = base.InsertParagraphBeforeSelf( text, trackChanges, formatting );
      p.PackagePart = this.PackagePart;

      this.NeedRefreshIndexes();

      return p;
    }

    public override void InsertPageBreakBeforeSelf()
    {
      base.InsertPageBreakBeforeSelf();
    }

    public override void InsertPageBreakAfterSelf()
    {
      base.InsertPageBreakAfterSelf();
    }

    [Obsolete( "Instead use: InsertHyperlink(Hyperlink h, int index)" )]
    public Paragraph InsertHyperlink( int index, Hyperlink h )
    {
      return InsertHyperlink( h, index );
    }


    public Paragraph InsertHyperlink( Hyperlink h, int index = 0 )
    {
      // Convert the path of this mainPart to its equilivant rels file path.
      var path = this.PackagePart.Uri.OriginalString.Replace( "/word/", "" );
      var rels_path = new Uri( String.Format( "/word/_rels/{0}.rels", path ), UriKind.Relative );

      // Check to see if the rels file exists and create it if not.
      if( !Document._package.PartExists( rels_path ) )
      {
        HelperFunctions.CreateRelsPackagePart( Document, rels_path );
      }

      // Check to see if a rel for this Picture exists, create it if not.
      var Id = HelperFunctions.GetOrGenerateRel( h.uri, this.PackagePart, TargetMode.External, HyperlinkRelation );

      XElement h_xml;
      if( index == 0 )
      {
        // Add this hyperlink as the first element.
        Xml.AddFirst( h.Xml );

        // Extract the picture back out of the DOM.
        h_xml = (XElement)Xml.FirstNode;
      }
      else
      {
        // Get the first run effected by this Insert
        Run run = GetFirstRunEffectedByEdit( index );

        if( run == null )
        {
          // Add this hyperlink as the last element.
          Xml.Add( h.Xml );

          // Extract the picture back out of the DOM.
          h_xml = (XElement)Xml.LastNode;
        }
        else
        {
          // Split this run at the point you want to insert
          XElement[] splitRun = Run.SplitRun( run, index );

          // Replace the origional run.
          run.Xml.ReplaceWith
          (
              splitRun[ 0 ],
              h.Xml,
              splitRun[ 1 ]
          );

          // Get the first run effected by this Insert
          run = GetFirstRunEffectedByEdit( index );

          // The picture has to be the next element, extract it back out of the DOM.
          h_xml = (XElement)run.Xml.NextNode;
        }
      }

      h_xml.SetAttributeValue( Document.r + "id", Id );

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();

      return this;
    }

    public void RemoveHyperlink( int index )
    {
      // Dosen't make sense to remove a Hyperlink at a negative index.
      if( index < 0 )
        throw new ArgumentOutOfRangeException();

      // Need somewhere to store the count.
      int count = 0;
      bool found = false;
      RemoveHyperlinkRecursive( Xml, index, ref count, ref found );

      // If !found then the user tried to remove a hyperlink at an index greater than the last. 
      if( !found )
        throw new ArgumentOutOfRangeException();
    }

    public override Paragraph InsertParagraphAfterSelf( Paragraph p )
    {
      this.ValidateInsert();
      this.InsertFollowingTables( p, true );
      this.ClearContainerParagraphsCache();

      var p2 = base.InsertParagraphAfterSelf( p );
      p2.PackagePart = this.PackagePart;

      return p2;
    }

    public override Paragraph InsertParagraphAfterSelf( string text, bool trackChanges, Formatting formatting )
    {
      this.ValidateInsert();
      this.ClearContainerParagraphsCache();

      var p = base.InsertParagraphAfterSelf( text, trackChanges, formatting );
      p.PackagePart = this.PackagePart;

      return p;
    }

    public override Paragraph InsertParagraphAfterSelf( string text, bool trackChanges )
    {
      this.ValidateInsert();
      this.ClearContainerParagraphsCache();

      var p = base.InsertParagraphAfterSelf( text, trackChanges );
      p.PackagePart = this.PackagePart;

      return p;
    }

    public override Paragraph InsertParagraphAfterSelf( string text )
    {
      this.ValidateInsert();
      this.ClearContainerParagraphsCache();

      var p = base.InsertParagraphAfterSelf( text );
      p.PackagePart = this.PackagePart;

      return p;
    }

    public void Remove( bool trackChanges, RemoveParagraphFlags removeParagraphFlags = RemoveParagraphFlags.All )
    {
      var mainContainer = this.GetMainParentContainer();
      if( mainContainer != null )
      {
        mainContainer.RemoveParagraph( this, trackChanges, removeParagraphFlags );
        mainContainer.NeedRefreshParagraphIndexes = true;
      }
    }

    //public Picture InsertPicture(Picture picture)
    //{
    //    Picture newPicture = picture;
    //    newPicture.i = new XElement(picture.i);

    //    xml.Add(newPicture.i);
    //    pictures.Add(newPicture);
    //    return newPicture;  
    //}

    // <summary>
    // Insert a Picture at the end of this paragraph.
    // </summary>
    // <param name="description">A string to describe this Picture.</param>
    // <param name="imageID">The unique id that identifies the Image this Picture represents.</param>
    // <param name="name">The name of this image.</param>
    // <returns>A Picture.</returns>
    // <example>
    // <code>
    // // Create a document using a relative filename.
    // using (var document = DocX.Create(@"Test.docx"))
    // {
    //     // Add a new Paragraph to this document.
    //     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
    //
    //     // Add an Image to this document.
    //     Xceed.Document.NET.Image img = document.AddImage(@"Image.jpg");
    //
    //     // Insert pic at the end of Paragraph p.
    //     Picture pic = p.InsertPicture(img.Id, "Photo 31415", "A pie I baked.");
    //
    //     // Rotate the Picture clockwise by 30 degrees. 
    //     pic.Rotation = 30;
    //
    //     // Resize the Picture.
    //     pic.Width = 400;
    //     pic.Height = 300;
    //
    //     // Set the shape of this Picture to be a cube.
    //     pic.SetPictureShape(BasicShapes.cube);
    //
    //     // Flip the Picture Horizontally.
    //     pic.FlipHorizontal = true;
    //
    //     // Save all changes made to this document.
    //     document.Save();
    // }// Release this document from memory.
    // </code>
    // </example>
    // Removed to simplify the API.
    //public Picture InsertPicture(string imageID, string name, string description)
    //{
    //    Picture p = CreatePicture(Document, imageID, name, description);
    //    Xml.Add(p.Xml);
    //    return p;
    //}

    // Removed because it confusses the API.
    //public Picture InsertPicture(string imageID)
    //{
    //    return InsertPicture(imageID, string.Empty, string.Empty);
    //}

    //public Picture InsertPicture(int index, Picture picture)
    //{
    //    Picture p = picture;
    //    p.i = new XElement(picture.i);

    //    Run run = GetFirstRunEffectedByEdit(index);

    //    if (run == null)
    //        xml.Add(p.i);
    //    else
    //    {
    //        // Split this run at the point you want to insert
    //        XElement[] splitRun = Run.SplitRun(run, index);

    //        // Replace the origional run
    //        run.Xml.ReplaceWith
    //        (
    //            splitRun[0],
    //            p.i,
    //            splitRun[1]
    //        );
    //    }

    //    // Rebuild the run lookup for this paragraph
    //    runLookup.Clear();
    //    BuildRunLookup(xml);
    //    Document.RenumberIDs(document);
    //    return p;
    //}

    // <summary>
    // Insert a Picture into this Paragraph at a specified index.
    // </summary>
    // <param name="description">A string to describe this Picture.</param>
    // <param name="imageID">The unique id that identifies the Image this Picture represents.</param>
    // <param name="name">The name of this image.</param>
    // <param name="index">The index to insert this Picture at.</param>
    // <returns>A Picture.</returns>
    // <example>
    // <code>
    // // Create a document using a relative filename.
    // using (var document = DocX.Create(@"Test.docx"))
    // {
    //     // Add a new Paragraph to this document.
    //     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
    //
    //     // Add an Image to this document.
    //     Xceed.Document.NET.Image img = document.AddImage(@"Image.jpg");
    //
    //     // Insert pic at the start of Paragraph p.
    //     Picture pic = p.InsertPicture(0, img.Id, "Photo 31415", "A pie I baked.");
    //
    //     // Rotate the Picture clockwise by 30 degrees. 
    //     pic.Rotation = 30;
    //
    //     // Resize the Picture.
    //     pic.Width = 400;
    //     pic.Height = 300;
    //
    //     // Set the shape of this Picture to be a cube.
    //     pic.SetPictureShape(BasicShapes.cube);
    //
    //     // Flip the Picture Horizontally.
    //     pic.FlipHorizontal = true;
    //
    //     // Save all changes made to this document.
    //     document.Save();
    // }// Release this document from memory.
    // </code>
    // </example>
    // Removed to simplify API.
    //public Picture InsertPicture(int index, string imageID, string name, string description)
    //{
    //    Picture picture = CreatePicture(Document, imageID, name, description);

    //    Run run = GetFirstRunEffectedByEdit(index);

    //    if (run == null)
    //        Xml.Add(picture.Xml);
    //    else
    //    {
    //        // Split this run at the point you want to insert
    //        XElement[] splitRun = Run.SplitRun(run, index);

    //        // Replace the origional run
    //        run.Xml.ReplaceWith
    //        (
    //            splitRun[0],
    //            picture.Xml,
    //            splitRun[1]
    //        );
    //    }

    //    HelperFunctions.RenumberIDs(Document);
    //    return picture;
    //}

    // Removed because it confusses the API.
    //public Picture InsertPicture(int index, string imageID)
    //{
    //    return InsertPicture(index, imageID, string.Empty, string.Empty);
    //}

    public void InsertText( string value, bool trackChanges = false, Formatting formatting = null )
    {
      this.InsertText( this.Text.Length, value, trackChanges, formatting );
    }

    public void InsertText( int index, string value, bool trackChanges = false, Formatting formatting = null )
    {
      // Timestamp to mark the start of insert
      var now = DateTime.Now;
      var insert_datetime = new DateTime( now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc );

      // Get the first run effected by this Insert
      var run = this.GetFirstRunEffectedByEdit( index );

      if( run == null )
      {
        object insert = ( formatting != null ) ? HelperFunctions.FormatInput( value, formatting.Xml ) : HelperFunctions.FormatInput( value, null );

        if( trackChanges )
        {
          insert = Document.CreateEdit( EditType.ins, insert_datetime, insert );
        }
        this.Xml.Add( insert );
      }
      else
      {
        object newRuns = null;
        var rPr = run.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );

        if( formatting != null )
        {
          Formatting oldFormatting = null;
          Formatting newFormatting = null;

          if( rPr != null )
          {
            oldFormatting = Formatting.Parse( rPr, null, null, this.Document );
            if( oldFormatting != null )
            {
              // Clone formatting and apply received formatting 
              newFormatting = oldFormatting.Clone();
              this.ApplyFormattingFrom( ref newFormatting, formatting );
            }
            else
            {
              newFormatting = formatting;
            }
          }
          else
          {
            newFormatting = formatting;
          }

          newRuns = HelperFunctions.FormatInput( value, newFormatting.Xml );
        }
        else
        {
          newRuns = HelperFunctions.FormatInput( value, rPr );
        }

        // The parent of this Run
        var parentElement = run.Xml.Parent;
        switch( parentElement.Name.LocalName )
        {
          case "ins":
            {
              // The datetime that this ins was created
              var parent_ins_date = DateTime.Parse( parentElement.Attribute( XName.Get( "date", Document.w.NamespaceName ) ).Value );

              /* 
               * Special case: You want to track changes,
               * and the first Run effected by this insert
               * has a datetime stamp equal to now.
              */
              if( trackChanges && parent_ins_date.CompareTo( insert_datetime ) == 0 )
              {
                /*
                 * Inserting into a non edit and this special case, is the same procedure.
                */
                goto default;
              }

              /*
               * If not the special case above, 
               * then inserting into an ins or a del, is the same procedure.
              */
              goto case "del";
            }

          case "del":
            {
              object insert = newRuns;
              if( trackChanges )
              {
                insert = Document.CreateEdit( EditType.ins, insert_datetime, newRuns );
              }

              // Split this Edit at the point you want to insert
              var splitEdit = SplitEdit( parentElement, index, EditType.ins );

              // Replace the origional run
              parentElement.ReplaceWith
              (
                  splitEdit[ 0 ],
                  insert,
                  splitEdit[ 1 ]
              );

              break;
            }

          default:
            {
              object insert = newRuns;
              if( trackChanges && !parentElement.Name.LocalName.Equals( "ins" ) )
              {
                insert = Document.CreateEdit( EditType.ins, insert_datetime, newRuns );
              }
              // Split this run at the point you want to insert
              var splitRun = Run.SplitRun( run, index );

              // Replace the origional run
              run.Xml.ReplaceWith
              (
                  splitRun[ 0 ],
                  insert,
                  splitRun[ 1 ]
              );

              break;
            }
        }
      }

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();
    }

    public Paragraph Culture( CultureInfo culture )
    {
      this.ApplyTextFormattingProperty( XName.Get( "lang", Document.w.NamespaceName ),
                                        string.Empty,
                                        new XAttribute( XName.Get( "val", Document.w.NamespaceName ), culture.Name ) );

      return this;
    }

    public Paragraph Append( string text )
    {
      var newRuns = HelperFunctions.FormatInput( text, null );
      this.Xml.Add( newRuns );

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).Reverse().Take( newRuns.Count() ).ToList();

      this.NeedRefreshIndexes();

      return this;
    }

    public Paragraph Append( string text, Formatting format )
    {
      // Text
      this.Append( text );

      // Bold
      if( format.Bold.HasValue && format.Bold.Value )
        Bold();

      // CapsStyle
      if( format.CapsStyle.HasValue )
        CapsStyle( format.CapsStyle.Value );

      // FontColor
      if( format.FontColor.HasValue )
        Color( format.FontColor.Value );

      // FontFamily
      if( format.FontFamily != null )
        Font( format.FontFamily );

      // Hidden
      if( format.Hidden.HasValue && format.Hidden.Value )
        Hide();

      // Highlight
      if( format.Highlight.HasValue )
        Highlight( format.Highlight.Value );

      // Shading
      if( format.Shading.HasValue )
        Shading( format.Shading.Value );

      // ShadingPattern
      if( format.ShadingPattern != null )
        ShadingPattern( format.ShadingPattern );

      // Border
      if( format.Border != null )
        Border( format.Border );

      // Italic
      if( format.Italic.HasValue && format.Italic.Value )
        Italic();

      // Kerning
      if( format.Kerning.HasValue )
        Kerning( format.Kerning.Value );

      // Language
      if( format.Language != null )
        Culture( format.Language );

      // Misc
      if( format.Misc.HasValue )
        Misc( format.Misc.Value );

      // PercentageScale
      if( format.PercentageScale.HasValue )
        PercentageScale( format.PercentageScale.Value );

      // Position
      if( format.Position.HasValue )
        Position( format.Position.Value );

      // Script
      if( format.Script.HasValue )
        Script( format.Script.Value );

      // Size
      if( format.Size.HasValue )
        FontSize( format.Size.Value );

      // Spacing
      if( format.Spacing.HasValue )
        Spacing( format.Spacing.Value );

      // StrikeThrough
      if( format.StrikeThrough.HasValue )
        StrikeThrough( format.StrikeThrough.Value );

      // UnderlineColor
      if( format.UnderlineColor.HasValue )
        UnderlineColor( format.UnderlineColor.Value );

      // UnderlineStyle
      if( format.UnderlineStyle.HasValue )
        UnderlineStyle( format.UnderlineStyle.Value );

      return this;
    }

    public Paragraph AppendHyperlink( Hyperlink h )
    {
      // Convert the path of this mainPart to its equilivant rels file path.
      var path = this.PackagePart.Uri.OriginalString.Replace( "/word/", "" );
      var rels_path = new Uri( "/word/_rels/" + path + ".rels", UriKind.Relative );

      // Check to see if the rels file exists and create it if not.
      if( !Document._package.PartExists( rels_path ) )
      {
        HelperFunctions.CreateRelsPackagePart( Document, rels_path );
      }

      // Check to see if a rel for this Hyperlink exists, create it if not.
      var Id = HelperFunctions.GetOrGenerateRel( h.uri, this.PackagePart, TargetMode.External, HyperlinkRelation );

      this.Xml.Add( h.Xml );
      this.Xml.Elements().Last().SetAttributeValue( Document.r + "id", Id );

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();
      return this;
    }

    public Paragraph AppendPicture( Picture p )
    {
      // Convert the path of this mainPart to its equilivant rels file path.
      var path = this.PackagePart.Uri.OriginalString.Replace( "/word/", "" );
      var rels_path = new Uri( "/word/_rels/" + path + ".rels", UriKind.Relative );

      // Check to see if the rels file exists and create it if not.
      if( !Document._package.PartExists( rels_path ) )
      {
        HelperFunctions.CreateRelsPackagePart( this.Document, rels_path );
      }

      // Check to see if a rel for this Picture exists, create it if not.
      var rel_Id = HelperFunctions.GetOrGenerateRel( p._img._pr.TargetUri, this.PackagePart, TargetMode.Internal, Document.RelationshipImage );

      // Add the Picture Xml to the end of the Paragragraph Xml.
      this.Xml.Add( p.Xml );

      // Extract the attribute id from the Pictures Xml.
      var embed_id =
      (
          from e in this.Xml.Elements().Last().Descendants()
          where e.Name.LocalName.Equals( "blip" )
          select e.Attribute( XName.Get( "embed", Document.r.NamespaceName ) )
      ).Single();

      // Set its value to the Pictures relationships id.
      embed_id.SetValue( rel_Id );

      // Extract the attribute id from the Pictures Xml.
      var docPr =
      (
          from e in this.Xml.Elements().Last().Descendants()
          where e.Name.LocalName.Equals( "docPr" )
          select e
      ).Single();

      // Set its value to a unique id.
      docPr.SetAttributeValue( "id", this.Document.GetNextFreeDocPrId().ToString() );

      // For formatting such as .Bold()
      // _runs = Xml.Elements( XName.Get( "r", Document.w.NamespaceName ) ).Reverse().Take( p.Xml.Elements( XName.Get( "r", Document.w.NamespaceName ) ).Count() ).ToList();
      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      return this;
    }


























    public Paragraph AppendNote( Note note, Formatting noteNumberFormatting = null )
    {
      if( note != null )
      {
        note.SetContainedParagraph( this );

        // Append the note to the paragraph and format the number in the paragraph.
        this.Xml.Add( note.CreateReferenceRun( noteNumberFormatting ) );
      }

      return this;
    }

    public Paragraph NextParagraph
    {
      get
      {
        if( this.Xml == null )
          return null;

        var paragraphs = this.GetContainerParagraphs();

        if( paragraphs == null )
          return null;

        var index = paragraphs.IndexOf( paragraphs.FirstOrDefault( p => p.Xml == this.Xml ) );
        if( ( index < 0 ) || ( index >= paragraphs.Count - 1 ) )
          return null;

        return paragraphs[ index + 1 ];
      }
    }

    public Paragraph PreviousParagraph
    {
      get
      {
        if( this.Xml == null )
          return null;

        var paragraphs = this.GetContainerParagraphs();

        if( paragraphs == null )
          return null;

        var index = paragraphs.IndexOf( paragraphs.FirstOrDefault( p => p.Xml == this.Xml ) );
        if( index < 1 )
          return null;

        return paragraphs[ index - 1 ];
      }
    }

    public Paragraph AppendEquation( String equation, Alignment align = Alignment.center )
    {
      var alignString = string.Empty;
      switch( align )
      {
        case Alignment.left:
          alignString = "left";
          break;
        case Alignment.right:
          alignString = "right";
          break;
        default:
          alignString = "center";
          break;
      }

      // Create equation element
      XElement oMathPara =
          new XElement
          (
              XName.Get( "oMathPara", Document.m.NamespaceName ),
              new XElement[]
              {
                new XElement
                (
                  XName.Get( "oMathParaPr", Document.m.NamespaceName ),
                  new XElement
                  (
                    XName.Get( "jc", Document.m.NamespaceName ),
                    new XAttribute( XName.Get( "val", Document.m.NamespaceName ), alignString )
                  )
                ),

                new XElement
                (
                  XName.Get( "oMath", Document.m.NamespaceName ),
                  new XElement
                  (
                      XName.Get( "r", Document.w.NamespaceName ),
                      new Formatting() { FontFamily = new Font( "Cambria Math" ) }.Xml,                           // create formatting
                      new XElement( XName.Get( "t", Document.m.NamespaceName ), equation )                            // create equation string
                  )
                )
              }
          );

      // Add equation element into paragraph xml and update runs collection
      this.Xml.Add( oMathPara );
      _runs = this.Xml.Elements( XName.Get( "oMathPara", Document.m.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();

      // Return paragraph with equation
      return this;
    }

    public Paragraph InsertPicture( Picture p, int index = 0 )
    {
      // Convert the path of this mainPart to its equilivant rels file path.
      var path = this.PackagePart.Uri.OriginalString.Replace( "/word/", "" );
      var rels_path = new Uri( "/word/_rels/" + path + ".rels", UriKind.Relative );

      // Check to see if the rels file exists and create it if not.
      if( !Document._package.PartExists( rels_path ) )
      {
        HelperFunctions.CreateRelsPackagePart( Document, rels_path );
      }

      XElement p_xml;
      if( index == 0 )
      {
        // Add this picture before the first run.
        var firstRun = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).FirstOrDefault();
        if( firstRun != null )
        {
          firstRun.AddBeforeSelf( p.Xml );

          // Extract the picture back out of the DOM.
          p_xml = (XElement)firstRun.PreviousNode;
        }
        else
        {
          var pPr = this.Xml.Element( XName.Get( "pPr", Document.w.NamespaceName ) );
          if( pPr != null )
          {
            pPr.AddAfterSelf( p.Xml );
            p_xml = (XElement)pPr.NextNode;
          }
          else
          {
            this.Xml.AddFirst( p.Xml );
            p_xml = (XElement)this.Xml.FirstNode;
          }
        }
      }
      else
      {
        // Get the first run effected by this Insert
        var run = this.GetFirstRunEffectedByEdit( index );
        if( run == null )
        {
          // Add this picture as the last element.
          this.Xml.Add( p.Xml );

          // Extract the picture back out of the DOM.
          p_xml = (XElement)this.Xml.LastNode;
        }
        else
        {
          // Split this run at the point you want to insert
          var splitRun = Run.SplitRun( run, index );
          // Replace the original run.
          run.Xml.ReplaceWith( splitRun[ 0 ], p.Xml, splitRun[ 1 ] );

          // Get the first run effected by this Insert
          run = GetFirstRunEffectedByEdit( index );

          // The picture has to be the next element, extract it back out of the DOM.
          p_xml = (XElement)run.Xml.NextNode;
        }
      }

      // Extract the id attribute from the Pictures Xml.
      var embed_id =
      (
          from e in p_xml.Descendants()
          where e.Name.LocalName.Equals( "blip" )
          select e.Attribute( XName.Get( "embed", Document.r.NamespaceName ) )
      ).Single();

      var rel_id = HelperFunctions.GetOrGenerateRel( p._img._pr.TargetUri, this.PackagePart, TargetMode.Internal, Document.RelationshipImage );

      // Set its value to the Pictures relationships id.
      embed_id.SetValue( rel_id );

      // Extract the attribute id from the Pictures Xml.
      var docPr =
      (
          from e in p_xml.Descendants()
          where e.Name.LocalName.Equals( "docPr" )
          select e
      ).Single();

      // Set its value to a unique id.
      docPr.SetAttributeValue( "id", this.Document.GetNextFreeDocPrId().ToString() );

      return this;
    }





    public Paragraph InsertTabStopPosition( Alignment alignment, float position, TabStopPositionLeader leader = TabStopPositionLeader.none, int index = -1 )
    {
      var pPr = GetOrCreate_pPr();
      var tabs = pPr.Element( XName.Get( "tabs", Document.w.NamespaceName ) );
      if( tabs == null )
      {
        tabs = new XElement( XName.Get( "tabs", Document.w.NamespaceName ) );
        pPr.Add( tabs );
      }

      var newTab = new XElement( XName.Get( "tab", Document.w.NamespaceName ) );

      // Alignement
      var alignmentString = string.Empty;
      switch( alignment )
      {
        case Alignment.left:
          alignmentString = "left";
          break;
        case Alignment.right:
          alignmentString = "right";
          break;
        case Alignment.center:
          alignmentString = "center";
          break;
        default:
          throw new ArgumentException( "alignment", "Value must be left, right or center." );
      }
      newTab.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), alignmentString );

      // Position
      var posValue = position * 20.0f;
      newTab.SetAttributeValue( XName.Get( "pos", Document.w.NamespaceName ), posValue.ToString( CultureInfo.InvariantCulture ) );

      //Leader
      var leaderString = string.Empty;
      switch( leader )
      {
        case TabStopPositionLeader.none:
          leaderString = "none";
          break;
        case TabStopPositionLeader.dot:
          leaderString = "dot";
          break;
        case TabStopPositionLeader.underscore:
          leaderString = "underscore";
          break;
        case TabStopPositionLeader.hyphen:
          leaderString = "hyphen";
          break;
        default:
          throw new ArgumentException( "leader", "Unknown leader character." );
      }
      newTab.SetAttributeValue( XName.Get( "leader", Document.w.NamespaceName ), leaderString );

      var tabsList = tabs.Elements().ToList();
      if( ( index >= 0 ) && ( index < tabsList.Count() ) )
      {
        tabsList[ index ].AddBeforeSelf( newTab );
      }
      else
      {
        tabs.Add( newTab );
      }

      return this;
    }

    public Paragraph AppendLine( string text )
    {
      return Append( "\n" + text );
    }

    public Paragraph AppendLine()
    {
      return Append( "\n" );
    }

    public Paragraph Bold( bool isBold = true )
    {
      ApplyTextFormattingProperty( XName.Get( "b", Document.w.NamespaceName ), string.Empty, isBold ? null : new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "0" ) );

      return this;
    }

    public Paragraph Italic( bool isItalic = true )
    {
      ApplyTextFormattingProperty( XName.Get( "i", Document.w.NamespaceName ), string.Empty, isItalic ? null : new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "0" ) );

      return this;
    }

    public Paragraph Color( Xceed.Drawing.Color c )
    {
      ApplyTextFormattingProperty( XName.Get( "color", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), c.ToHex() ) );

      return this;
    }

    public Paragraph UnderlineStyle( UnderlineStyle underlineStyle )
    {
      string value;
      switch( underlineStyle )
      {
        case Xceed.Document.NET.UnderlineStyle.none:
          value = string.Empty;
          break;
        case Xceed.Document.NET.UnderlineStyle.singleLine:
          value = "single";
          break;
        case Xceed.Document.NET.UnderlineStyle.doubleLine:
          value = "double";
          break;
        default:
          value = underlineStyle.ToString();
          break;
      }

      ApplyTextFormattingProperty( XName.Get( "u", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), value ) );

      return this;
    }

    public Paragraph FontSize( double fontSize )
    {
      double tempSize = fontSize * 2;
      if( tempSize - (int)tempSize == 0 )
      {
        if( !( fontSize > 0 && fontSize < 1639 ) )
          throw new ArgumentException( "Size", "Value must be in the range 1 - 1638" );
      }

      else
        throw new ArgumentException( "Size", "Value must be either a whole or half number, examples: 32, 32.5" );

      ApplyTextFormattingProperty( XName.Get( "sz", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), fontSize * 2 ) );
      ApplyTextFormattingProperty( XName.Get( "szCs", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), fontSize * 2 ) );

      return this;
    }

    public Paragraph Font( string fontName )
    {
      return Font( new Font( fontName ) );
    }

    public Paragraph Font( Font fontFamily )
    {
      ApplyTextFormattingProperty
      (
          XName.Get( "rFonts", Document.w.NamespaceName ),
          string.Empty,
          new[]
          {
            new XAttribute(XName.Get("ascii", Document.w.NamespaceName), fontFamily.Name),
            new XAttribute(XName.Get("hAnsi", Document.w.NamespaceName), fontFamily.Name),
            new XAttribute(XName.Get("cs", Document.w.NamespaceName), fontFamily.Name),
            new XAttribute(XName.Get("eastAsia", Document.w.NamespaceName), fontFamily.Name),
          }
      );

      return this;
    }

    public Paragraph CapsStyle( CapsStyle capsStyle )
    {
      switch( capsStyle )
      {
        case Xceed.Document.NET.CapsStyle.none:
          break;

        default:
          {
            ApplyTextFormattingProperty( XName.Get( capsStyle.ToString(), Document.w.NamespaceName ), string.Empty, null );
            break;
          }
      }

      return this;
    }

    public Paragraph Script( Script script )
    {
      switch( script )
      {
        case Xceed.Document.NET.Script.none:
          break;

        default:
          {
            ApplyTextFormattingProperty( XName.Get( "vertAlign", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), script.ToString() ) );
            break;
          }
      }

      return this;
    }

    public Paragraph Highlight( Highlight highlight )
    {
      ApplyTextFormattingProperty( XName.Get( "highlight", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), highlight.ToString() ) );

      return this;
    }

    [Obsolete( "This method is obsolete and should no longer be used. Use the ShadingPattern method instead." )]
    public Paragraph Shading( Xceed.Drawing.Color shading, ShadingType shadingType = ShadingType.Text )
    {
      // Add to run
      if( shadingType == ShadingType.Text )
      {
        this.ApplyTextFormattingProperty( XName.Get( "shd", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "fill", Document.w.NamespaceName ), shading.ToHex() ) );
      }
      // Add to paragraph
      else
      {
        var pPr = GetOrCreate_pPr();
        var shd = pPr.Element( XName.Get( "shd", Document.w.NamespaceName ) );
        if( shd == null )
        {
          shd = new XElement( XName.Get( "shd", Document.w.NamespaceName ) );
          pPr.Add( shd );
        }

        var fillAttribute = shd.Attribute( XName.Get( "fill", Document.w.NamespaceName ) );
        if( fillAttribute == null )
        {
          shd.SetAttributeValue( XName.Get( "fill", Document.w.NamespaceName ), shading.ToHex() );
        }
        else
        {
          fillAttribute.SetValue( shading.ToHex() );
        }
      }

      return this;
    }

    public Paragraph ShadingPattern( ShadingPattern shadingPattern, ShadingType shadingType = ShadingType.Text )
    {
      // Add to run
      if( shadingType == ShadingType.Text )
      {
        this.ApplyTextFormattingProperty( XName.Get( "shd", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "fill", Document.w.NamespaceName ), shadingPattern.Fill.ToHex() ) );
        this.ApplyTextFormattingProperty( XName.Get( "shd", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), HelperFunctions.GetValueFromTablePatternStyle( shadingPattern.Style ) ) );
        this.ApplyTextFormattingProperty( XName.Get( "shd", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "color", Document.w.NamespaceName ), shadingPattern.StyleColor.ToHex() ) );
      }
      // Add to paragraph
      else
      {
        var pPr = GetOrCreate_pPr();
        var shd = pPr.Element( XName.Get( "shd", Document.w.NamespaceName ) );
        if( shd == null )
        {

          shd = new XElement( XName.Get( "shd", Document.w.NamespaceName ) );
          pPr.Add( shd );
        }

        shd.SetAttributeValue( XName.Get( "fill", Document.w.NamespaceName ), shadingPattern.Fill.ToHex() );
        shd.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), HelperFunctions.GetValueFromTablePatternStyle( shadingPattern.Style ) );
        shd.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), shadingPattern.Style == PatternStyle.Clear ? "auto" : shadingPattern.StyleColor.ToHex() );
      }

      return this;
    }

    public Paragraph Border( Border border )
    {
      var size = Xceed.Document.NET.Border.GetNumericSize( border.Size );

      var style = border.Tcbs.ToString().Remove( 0, 5 );

      this.ApplyTextFormattingProperty( XName.Get( "bdr", Document.w.NamespaceName ),
                                        string.Empty,
                                        new List<XAttribute>() { new XAttribute( XName.Get( "color", Document.w.NamespaceName ), border.Color.ToHex() ),
                                                                 new XAttribute( XName.Get( "space", Document.w.NamespaceName ), border.Space ),
                                                                 new XAttribute( XName.Get( "sz", Document.w.NamespaceName ), size ),
                                                                 new XAttribute( XName.Get( "val", Document.w.NamespaceName ), style ) } );

      return this;
    }

    public Paragraph Misc( Misc misc )
    {
      switch( misc )
      {
        case Xceed.Document.NET.Misc.none:
          break;

        case Xceed.Document.NET.Misc.outlineShadow:
          {
            ApplyTextFormattingProperty( XName.Get( "outline", Document.w.NamespaceName ), string.Empty, null );
            ApplyTextFormattingProperty( XName.Get( "shadow", Document.w.NamespaceName ), string.Empty, null );

            break;
          }

        case Xceed.Document.NET.Misc.engrave:
          {
            ApplyTextFormattingProperty( XName.Get( "imprint", Document.w.NamespaceName ), string.Empty, null );

            break;
          }

        default:
          {
            ApplyTextFormattingProperty( XName.Get( misc.ToString(), Document.w.NamespaceName ), string.Empty, null );

            break;
          }
      }

      return this;
    }

    public Paragraph StrikeThrough( StrikeThrough strikeThrough )
    {
      string value;
      switch( strikeThrough )
      {
        case Xceed.Document.NET.StrikeThrough.strike:
          value = "strike";
          break;
        case Xceed.Document.NET.StrikeThrough.doubleStrike:
          value = "dstrike";
          break;
        default:
          return this;
      }

      ApplyTextFormattingProperty( XName.Get( value, Document.w.NamespaceName ), string.Empty, null );

      return this;
    }

    public Paragraph UnderlineColor( Xceed.Drawing.Color underlineColor )
    {
      foreach( XElement run in _runs )
      {
        XElement rPr = run.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
        if( rPr == null )
        {
          run.AddFirst( new XElement( XName.Get( "rPr", Document.w.NamespaceName ) ) );
          rPr = run.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
        }

        XElement u = rPr.Element( XName.Get( "u", Document.w.NamespaceName ) );
        if( u == null )
        {
          rPr.SetElementValue( XName.Get( "u", Document.w.NamespaceName ), string.Empty );
          u = rPr.Element( XName.Get( "u", Document.w.NamespaceName ) );
          u.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), "single" );
        }

        u.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), underlineColor.ToHex() );
      }

      return this;
    }

    public Paragraph Hide()
    {
      ApplyTextFormattingProperty( XName.Get( "vanish", Document.w.NamespaceName ), string.Empty, null );

      return this;
    }

    public Paragraph Spacing( double spacing )
    {
      double tempSize = spacing * 20;
      if( tempSize - (int)tempSize == 0 )
      {
        if( !( spacing > -1585 && spacing < 1585 ) )
          throw new ArgumentException( "Spacing", "Value must be in the range: (-1584, 1584)" );
      }
      else
        throw new ArgumentException( "Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9" );

      ApplyTextFormattingProperty( XName.Get( "spacing", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), spacing * 20 ) );

      return this;
    }

    public Paragraph SpacingBefore( double spacingBefore )
    {
      spacingBefore *= 20;

      var pPr = GetOrCreate_pPr();
      var spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

      if( spacingBefore >= 0 )
      {
        if( spacing == null )
        {
          spacing = new XElement( XName.Get( "spacing", Document.w.NamespaceName ) );
          pPr.Add( spacing );
        }

        var beforeAttribute = spacing.Attribute( XName.Get( "before", Document.w.NamespaceName ) );

        if( beforeAttribute == null )
          spacing.SetAttributeValue( XName.Get( "before", Document.w.NamespaceName ), spacingBefore );
        else
          beforeAttribute.SetValue( spacingBefore );
      }

      if( Math.Abs( spacingBefore ) < 0f && spacing != null )
      {
        var beforeAttribute = spacing.Attribute( XName.Get( "before", Document.w.NamespaceName ) );
        if( beforeAttribute != null )
        {
          beforeAttribute.Remove();
        }

        if( !spacing.HasAttributes )
        {
          spacing.Remove();
        }
      }

      return this;
    }

    public Paragraph SpacingAfter( double spacingAfter )
    {
      spacingAfter *= 20;

      var pPr = GetOrCreate_pPr();
      var spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

      if( spacingAfter >= 0 )
      {
        if( spacing == null )
        {
          spacing = new XElement( XName.Get( "spacing", Document.w.NamespaceName ) );
          pPr.Add( spacing );
        }

        var afterAttribute = spacing.Attribute( XName.Get( "after", Document.w.NamespaceName ) );

        if( afterAttribute == null )
          spacing.SetAttributeValue( XName.Get( "after", Document.w.NamespaceName ), spacingAfter );
        else
          afterAttribute.SetValue( spacingAfter );
      }

      if( Math.Abs( spacingAfter ) < 0f && spacing != null )
      {
        var afterAttribute = spacing.Attribute( XName.Get( "after", Document.w.NamespaceName ) );
        if( afterAttribute != null )
        {
          afterAttribute.Remove();
        }

        if( !spacing.HasAttributes )
        {
          spacing.Remove();
        }
      }

      return this;
    }

    public Paragraph SpacingLine( double lineSpacing )
    {
      lineSpacing *= 20;

      var pPr = GetOrCreate_pPr();
      var spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

      if( lineSpacing >= 0 )
      {
        if( spacing == null )
        {
          spacing = new XElement( XName.Get( "spacing", Document.w.NamespaceName ) );
          pPr.Add( spacing );
        }

        var lineAttribute = spacing.Attribute( XName.Get( "line", Document.w.NamespaceName ) );

        if( lineAttribute == null )
          spacing.SetAttributeValue( XName.Get( "line", Document.w.NamespaceName ), lineSpacing );
        else
          lineAttribute.SetValue( lineSpacing );
      }

      if( Math.Abs( lineSpacing ) < 0f && spacing != null )
      {
        var lineAttribute = spacing.Attribute( XName.Get( "line", Document.w.NamespaceName ) );
        lineAttribute.Remove();

        if( !spacing.HasAttributes )
          spacing.Remove();
      }

      return this;
    }

    public Paragraph Kerning( float kerning )
    {
      ApplyTextFormattingProperty( XName.Get( "kern", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), kerning * 2f ) );

      return this;
    }

    public Paragraph Position( double position )
    {
      if( !( position > -1585 && position < 1585 ) )
        throw new ArgumentOutOfRangeException( "Position", "Value must be in the range -1585 - 1585" );

      ApplyTextFormattingProperty( XName.Get( "position", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), position * 2 ) );

      return this;
    }

    public Paragraph PercentageScale( float percentageScale )
    {
      if( !( percentageScale >= 1f && percentageScale <= 600f ) )
        throw new ArgumentException( "PercentageScale", "Value must be in the range: 1 - 600" );

      ApplyTextFormattingProperty( XName.Get( "w", Document.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", Document.w.NamespaceName ), percentageScale ) );

      return this;
    }

    public Paragraph AppendDocProperty( CustomProperty cp, bool trackChanges = false, Formatting f = null )
    {
      this.InsertDocProperty( cp, trackChanges, f );

      return this;
    }

    public DocProperty InsertDocProperty( CustomProperty cp, bool trackChanges = false, Formatting f = null )
    {
      XElement f_xml = null;
      if( f != null )
      {
        f_xml = f.Xml;
      }

      var e = new XElement( XName.Get( "fldSimple", Document.w.NamespaceName ),
                            new XAttribute( XName.Get( "instr", Document.w.NamespaceName ), string.Format( @"DOCPROPERTY {0} \* MERGEFORMAT", cp.Name ) ),
                            new XElement( XName.Get( "r", Document.w.NamespaceName ), new XElement( XName.Get( "t", Document.w.NamespaceName ), f_xml, cp.Value ) )
      );

      var xml = e;
      if( trackChanges )
      {
        var now = DateTime.Now;
        var insert_datetime = new DateTime( now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc );
        e = Document.CreateEdit( EditType.ins, insert_datetime, e );
      }

      this.Xml.Add( e );

      return new DocProperty( this.Document, xml );
    }

    public void RemoveText( int index, int count, bool trackChanges = false, bool removeEmptyParagraph = true )
    {
      if( count == 0 )
        return;

      // Timestamp to mark the start of insert
      var now = DateTime.Now;
      var remove_datetime = new DateTime( now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc );

      // The number of characters processed so far
      int processed = 0;

      do
      {
        // Get the first run effected by this Remove
        var run = GetFirstRunEffectedByEdit( index, EditType.del );

        // The parent of this Run
        var parentElement = run.Xml.Parent;
        switch( parentElement.Name.LocalName )
        {
          case "ins":
            {
              var splitEditBefore = this.SplitEdit( parentElement, index, EditType.del );
              var min = Math.Min( count - processed, run.Xml.ElementsAfterSelf().Sum( e => GetElementTextLength( e ) ) );
              var splitEditAfter = this.SplitEdit( parentElement, index + min, EditType.del );

              var temp = this.SplitEdit( splitEditBefore[ 1 ], index + min, EditType.del )[ 1 ];
              var middle = Document.CreateEdit( EditType.del, remove_datetime, temp.Elements() );
              processed += Paragraph.GetElementTextLength( middle as XElement );

              if( !trackChanges )
              {
                middle = null;
              }

              parentElement.ReplaceWith( splitEditBefore[ 0 ], middle, splitEditAfter[ 0 ] );

              processed += Paragraph.GetElementTextLength( middle as XElement );
              break;
            }

          case "del":
            {
              if( trackChanges )
              {
                // You cannot delete from a deletion, advance processed to the end of this del
                processed += Paragraph.GetElementTextLength( parentElement );
              }
              else
              {
                goto case "ins";
              }
              break;
            }

          default:
            {
              if( Paragraph.GetElementTextLength( run.Xml ) > 0 )
              {
                var splitRunBefore = Run.SplitRun( run, index, EditType.del );
                var min = Math.Min( index + ( count - processed ), run.EndIndex );
                var splitRunAfter = Run.SplitRun( run, min, EditType.del );

                var middle = Document.CreateEdit( EditType.del, remove_datetime, new List<XElement>() { Run.SplitRun( new Run( Document, splitRunBefore[ 1 ], run.StartIndex + GetElementTextLength( splitRunBefore[ 0 ] ) ), min, EditType.del )[ 0 ] } );
                processed = processed + Paragraph.GetElementTextLength( middle as XElement );

                if( !trackChanges )
                {
                  middle = null;
                }

                run.Xml.ReplaceWith( splitRunBefore[ 0 ], middle, splitRunAfter[ 1 ] );
              }
              else
              {
                processed = count;
              }
              break;
            }
        }

        // In some cases, removing an empty paragraph is allowed
        var canRemove = removeEmptyParagraph && GetElementTextLength( parentElement ) == 1 && string.IsNullOrEmpty( Text );

        if( parentElement.Parent != null )
        {
          // Need to make sure there is another paragraph in the parent cell
          if( parentElement.Parent.Name.LocalName == "tc" )
          {
            canRemove &= parentElement.Parent.Elements( XName.Get( "p", Document.w.NamespaceName ) ).Count() > 1;
          }

          // Need to make sure there is no drawing element within the parent element.
          // Picture elements contain no text length but they are still content.
          canRemove &= parentElement.Descendants( XName.Get( "drawing", Document.w.NamespaceName ) ).Count() == 0;

          if( canRemove )
          {
            parentElement.Remove();
            m_removed = true;
          }
        }
      }
      while( processed < count );

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();
    }

    public void RemoveText( int index, bool trackChanges = false )
    {
      this.RemoveText( index, Text.Length - index, trackChanges );
    }

    [Obsolete( "ReplaceText() with many parameters is obsolete. Use ReplaceText() with a StringReplaceTextOptions parameter instead." )]
    public void ReplaceText( string searchValue,
                             string newValue,
                             bool trackChanges = false,
                             RegexOptions options = RegexOptions.None,
                             Formatting newFormatting = null,
                             Formatting matchFormatting = null,
                             MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch,
                             bool escapeRegEx = true,
                             bool useRegExSubstitutions = false,
                             bool removeEmptyParagraph = true )
    {
      var mc = Regex.Matches( this.Text, escapeRegEx ? Regex.Escape( searchValue ) : searchValue, options );

      // Loop through the matches in reverse order
      foreach( Match m in mc.Cast<Match>().Reverse() )
      {
        // Assume the formatting matches until proven otherwise.
        bool formattingMatch = true;

        // Does the user want to match formatting?
        if( matchFormatting != null )
        {
          // The number of characters processed so far
          int processed = 0;

          do
          {
            // Get the next run effected
            var run = GetFirstRunEffectedByEdit( m.Index + processed );

            // Get this runs properties
            var rPr = run.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );

            if( rPr == null )
            {
              rPr = new Formatting().Xml;
            }

            /* 
             * Make sure that every formatting element in f.xml is also in this run,
             * if this is not true, then their formatting does not match.
             */
            if( !HelperFunctions.ContainsEveryChildOf( matchFormatting.Xml, rPr, fo ) )
            {
              formattingMatch = false;
              break;
            }

            // We have processed some characters, so update the counter.
            processed += run.Value.Length;

          } while( processed < m.Length );
        }

        // If the formatting matches, do the replace.
        if( formattingMatch )
        {
          //perform RegEx substitutions. Only named groups are not supported. Everything else is supported. However character escapes are not covered.
          if( useRegExSubstitutions && !string.IsNullOrEmpty( newValue ) )
          {
            newValue = newValue.Replace( "$&", m.Value );
            if( m.Groups.Count > 0 )
            {
              int lastcap = 0;
              for( int k = 0; k < m.Groups.Count; k++ )
              {
                var g = m.Groups[ k ];
                if( ( g == null ) || ( g.Value == "" ) )
                  continue;
                newValue = newValue.Replace( "$" + k.ToString(), g.Value );
                lastcap = k;
              }
              newValue = newValue.Replace( "$+", m.Groups[ lastcap ].Value );
            }
            if( m.Index > 0 )
            {
              newValue = newValue.Replace( "$`", this.Text.Substring( 0, m.Index ) );
            }
            if( ( m.Index + m.Length ) < this.Text.Length )
            {
              newValue = newValue.Replace( "$'", this.Text.Substring( m.Index + m.Length ) );
            }
            newValue = newValue.Replace( "$_", this.Text );
            newValue = newValue.Replace( "$$", "$" );
          }

          if( !string.IsNullOrEmpty( newValue ) )
          {
            this.InsertText( m.Index + m.Length, newValue, trackChanges, newFormatting );
          }
          if( m.Length > 0 )
          {
            this.RemoveText( m.Index, m.Length, trackChanges, removeEmptyParagraph );
          }
        }
      }
    }

    [Obsolete( "ReplaceText() with many parameters is obsolete. Use ReplaceText() with a FunctionReplaceTextOptions parameter instead." )]
    public void ReplaceText( string findPattern, Func<string, string> regexMatchHandler, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch, bool removeEmptyParagraph = true )
    {
      var matchCol = Regex.Matches( this.Text, findPattern, options );
      var reversedMatchCol = matchCol.Cast<Match>().Reverse();

      foreach( var match in reversedMatchCol )
      {
        var formattingMatch = true;

        if( matchFormatting != null )
        {
          int processed = 0;

          while( processed < match.Length )
          {
            var run = this.GetFirstRunEffectedByEdit( match.Index + processed );
            var rPr = run.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
            if( rPr == null )
            {
              rPr = new Formatting().Xml;
            }

            // Make sure that every formatting element in matchFormatting.Xml is also in this run,
            // if false => formatting does not match.
            if( !HelperFunctions.ContainsEveryChildOf( matchFormatting.Xml, rPr, fo ) )
            {
              formattingMatch = false;
              break;
            }

            processed += run.Value.Length;
          }
        }

        // Replace text when formatting matches.
        if( formattingMatch )
        {
          int lastcap = 0;
          for( int k = 0; k < match.Groups.Count; k++ )
          {
            var g = match.Groups[ k ];
            if( ( g == null ) || ( g.Value == "" ) )
              continue;
            lastcap = k;
          }

          var newValue = regexMatchHandler.Invoke( match.Groups[ lastcap ].Value );
          this.InsertText( match.Index + match.Value.Length, newValue, trackChanges, newFormatting );
          this.RemoveText( match.Index, match.Value.Length, trackChanges, removeEmptyParagraph );
        }
      }
    }

    [Obsolete( "ReplaceText() with many parameters is obsolete. Use ReplaceText() with an ObjectReplaceTextOptions parameter instead." )]
    public void ReplaceTextWithObject( string searchValue,
                                       DocumentElement objectToAdd,
                                       bool trackChanges = false,
                                       RegexOptions options = RegexOptions.None,
                                       Formatting matchFormatting = null,
                                       MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch,
                                       bool escapeRegEx = true,
                                       bool removeEmptyParagraph = true )
    {
      var mc = Regex.Matches( this.Text, escapeRegEx ? Regex.Escape( searchValue ) : searchValue, options );

      // Loop through the matches in reverse order
      foreach( Match m in mc.Cast<Match>().Reverse() )
      {
        // Assume the formatting matches until proven otherwise.
        bool formattingMatch = true;

        // Does the user want to match formatting?
        if( matchFormatting != null )
        {
          // The number of characters processed so far
          int processed = 0;

          do
          {
            // Get the next run effected
            var run = GetFirstRunEffectedByEdit( m.Index + processed );

            // Get this runs properties
            var rPr = run.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );

            if( rPr == null )
            {
              rPr = new Formatting().Xml;
            }

            /* 
             * Make sure that every formatting element in f.xml is also in this run,
             * if this is not true, then their formatting does not match.
             */
            if( !HelperFunctions.ContainsEveryChildOf( matchFormatting.Xml, rPr, fo ) )
            {
              formattingMatch = false;
              break;
            }

            // We have processed some characters, so update the counter.
            processed += run.Value.Length;

          } while( processed < m.Length );
        }

        // If the formatting matches, do the replace.
        if( formattingMatch )
        {
          if( objectToAdd != null )
          {
            if( objectToAdd is Picture )
            {
              this.InsertPicture( (Picture)objectToAdd, m.Index + m.Length );
            }
            else if( objectToAdd is Hyperlink )
            {
              this.InsertHyperlink( (Hyperlink)objectToAdd, m.Index + m.Length );
            }
            else if( objectToAdd is Table )
            {
              this.InsertTableAfterSelf( (Table)objectToAdd );
            }
            else
            {
              throw new ArgumentException( "Unknown object received. Valid objects are Picture, Hyperlink or Table." );
            }
          }
          if( m.Length > 0 )
          {
            this.RemoveText( m.Index, m.Length, trackChanges, removeEmptyParagraph );
          }
        }
      }
    }

    public bool ReplaceText( StringReplaceTextOptions replaceTextOptions )
    {
      if( string.IsNullOrEmpty( replaceTextOptions.SearchValue ) )
        throw new ArgumentException( "searchValue cannot be null or empty.", "searchValue" );
      if( replaceTextOptions.NewValue == null )
        throw new ArgumentException( "newValue cannot be null.", "newValue" );

      var replaceSuccess = false;
      var textToAnalyse = this.Text;

      if( !this.UpdateTextToReplace( replaceTextOptions, ref textToAnalyse ) )
        return replaceSuccess;

      if( replaceTextOptions.StopAfterOneReplacement )
      {
        var singleMatch = Regex.Match( textToAnalyse, replaceTextOptions.EscapeRegEx ? Regex.Escape( replaceTextOptions.SearchValue ) : replaceTextOptions.SearchValue, replaceTextOptions.RegExOptions );
        if( singleMatch.Length > 0 )
        {
          replaceSuccess = this.ReplaceTextCore( singleMatch, replaceTextOptions );
        }
      }
      else
      {
        var mc = Regex.Matches( textToAnalyse, replaceTextOptions.EscapeRegEx ? Regex.Escape( replaceTextOptions.SearchValue ) : replaceTextOptions.SearchValue, replaceTextOptions.RegExOptions );

        // Loop through the matches in reverse order
        foreach( Match match in mc.Cast<Match>().Reverse() )
        {
          var result = this.ReplaceTextCore( match, replaceTextOptions );
          if( !replaceSuccess )
          {
            replaceSuccess = result;
          }
        }
      }
      return replaceSuccess;
    }






















    public bool ReplaceText( FunctionReplaceTextOptions replaceTextOptions )
    {
      if( string.IsNullOrEmpty( replaceTextOptions.FindPattern ) )
        throw new ArgumentException( "FindPattern cannot be null or empty.", "FindPattern" );
      if( replaceTextOptions.RegexMatchHandler == null )
        throw new ArgumentException( "RegexMatchHandler cannot be null.", "RegexMatchHandler" );

      var replaceSuccess = false;
      var textToAnalyse = this.Text;

      if( !this.UpdateTextToReplace( replaceTextOptions, ref textToAnalyse ) )
        return replaceSuccess;

      if( replaceTextOptions.StopAfterOneReplacement )
      {
        var singleMatch = Regex.Match( textToAnalyse, replaceTextOptions.FindPattern, replaceTextOptions.RegExOptions );
        if( singleMatch.Length > 0 )
        {
          replaceSuccess = this.ReplaceTextCore( singleMatch, replaceTextOptions );
        }
      }
      else
      {
        var matchCol = Regex.Matches( textToAnalyse, replaceTextOptions.FindPattern, replaceTextOptions.RegExOptions );
        var reversedMatchCol = matchCol.Cast<Match>().Reverse();

        foreach( var match in reversedMatchCol )
        {
          var result = this.ReplaceTextCore( match, replaceTextOptions );
          if( !replaceSuccess )
          {
            replaceSuccess = result;
          }
        }
      }
      return replaceSuccess;
    }

    public bool ReplaceTextWithObject( ObjectReplaceTextOptions replaceTextOptions )
    {
      if( string.IsNullOrEmpty( replaceTextOptions.SearchValue ) )
        throw new ArgumentException( "searchValue cannot be null or empty.", "searchValue" );
      if( replaceTextOptions.NewObject == null )
        throw new ArgumentException( "NewObject cannot be null.", "newValue" );

      var replaceSuccess = false;
      var textToAnalyse = this.Text;

      if( !this.UpdateTextToReplace( replaceTextOptions, ref textToAnalyse ) )
        return replaceSuccess;

      if( replaceTextOptions.StopAfterOneReplacement )
      {
        var singleMatch = Regex.Match( textToAnalyse, replaceTextOptions.EscapeRegEx ? Regex.Escape( replaceTextOptions.SearchValue ) : replaceTextOptions.SearchValue, replaceTextOptions.RegExOptions );
        if( singleMatch.Length > 0 )
        {
          replaceSuccess = this.ReplaceTextCore( singleMatch, replaceTextOptions );
        }
      }
      else
      {
        var mc = Regex.Matches( textToAnalyse, replaceTextOptions.EscapeRegEx ? Regex.Escape( replaceTextOptions.SearchValue ) : replaceTextOptions.SearchValue, replaceTextOptions.RegExOptions );

        // Loop through the matches in reverse order
        foreach( Match m in mc.Cast<Match>().Reverse() )
        {
          var result = this.ReplaceTextCore( m, replaceTextOptions );
          if( !replaceSuccess )
          {
            replaceSuccess = result;
          }
        }
      }

      return replaceSuccess;
    }

    public List<int> FindAll( string str )
    {
      return this.FindAll( str, RegexOptions.None );
    }

    public List<int> FindAll( string str, RegexOptions options )
    {
      var mc = Regex.Matches( this.Text, Regex.Escape( str ), options );

      var query =
      (
          from m in mc.Cast<Match>()
          select m.Index
      ).ToList();

      return query;
    }

    public List<string> FindAllByPattern( string str, RegexOptions options )
    {
      MatchCollection mc = Regex.Matches( this.Text, str, options );

      var query =
      (
          from m in mc.Cast<Match>()
          select m.Value
      ).ToList();

      return query;
    }

    public void InsertPageNumber( PageNumberFormat? pnf = null, int index = 0 )
    {
      var fldSimple = this.SetPageNumberFields( pnf );
      var content = this.GetNumberContentBasedOnLast_rPr();

      fldSimple.Add( content );

      if( index == 0 )
      {
        Xml.AddFirst( fldSimple );
      }
      else
      {
        var r = GetFirstRunEffectedByEdit( index, EditType.ins );
        var splitEdit = SplitEdit( r.Xml, index, EditType.ins );
        r.Xml.ReplaceWith
        (
            splitEdit[ 0 ],
            fldSimple,
            splitEdit[ 1 ]
        );
      }

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();
    }

    public Paragraph AppendPageNumber( PageNumberFormat? pnf = null )
    {
      var fldSimple = this.SetPageNumberFields( pnf );
      var content = this.GetNumberContentBasedOnLast_rPr();

      fldSimple.Add( content );
      Xml.Add( fldSimple );

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();
      return this;
    }

    public void InsertPageCount( PageNumberFormat? pnf = null, int index = 0, bool useSectionPageCount = true )
    {
      var fldSimple = this.SetPageCountFields( pnf, useSectionPageCount );
      var content = this.GetNumberContentBasedOnLast_rPr();

      fldSimple.Add( content );

      if( index == 0 )
        Xml.AddFirst( fldSimple );
      else
      {
        Run r = GetFirstRunEffectedByEdit( index, EditType.ins );
        XElement[] splitEdit = SplitEdit( r.Xml, index, EditType.ins );
        r.Xml.ReplaceWith
        (
            splitEdit[ 0 ],
            fldSimple,
            splitEdit[ 1 ]
        );
      }

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();
    }

    public Paragraph AppendPageCount( PageNumberFormat? pnf = null, bool useSectionPageCount = false )
    {
      var fldSimple = this.SetPageCountFields( pnf, useSectionPageCount );
      var content = this.GetNumberContentBasedOnLast_rPr();

      fldSimple.Add( content );
      Xml.Add( fldSimple );

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();

      this.NeedRefreshIndexes();
      return this;
    }


    public void SetLineSpacing( LineSpacingType spacingType, float spacingFloat )
    {
      var pPr = this.GetOrCreate_pPr();
      var spacingXName = XName.Get( "spacing", Document.w.NamespaceName );
      var spacing = pPr.Element( spacingXName );
      if( spacing == null )
      {
        pPr.Add( new XElement( spacingXName ) );
        spacing = pPr.Element( spacingXName );
      }

      var spacingTypeAttribute = ( spacingType == LineSpacingType.Before )
                                 ? "before"
                                 : ( spacingType == LineSpacingType.After ) ? "after" : "line";
      spacing.SetAttributeValue( XName.Get( spacingTypeAttribute, Document.w.NamespaceName ), (int)( spacingFloat * 20f ) );
    }

    public void SetLineSpacing( LineSpacingTypeAuto spacingTypeAuto )
    {
      var pPr = this.GetOrCreate_pPr();
      var spacingXName = XName.Get( "spacing", Document.w.NamespaceName );
      var spacing = pPr.Element( spacingXName );

      if( spacingTypeAuto == LineSpacingTypeAuto.None )
      {
        if( spacing != null )
        {
          spacing.Remove();
        }
      }
      else
      {
        if( spacing == null )
        {
          pPr.Add( new XElement( spacingXName ) );
          spacing = pPr.Element( spacingXName );
        }

        int spacingValue = 500;
        var spacingTypeAttribute = ( spacingTypeAuto == LineSpacingTypeAuto.AutoAfter ) ? "after" : "before";
        var autoSpacingTypeAttribute = ( spacingTypeAuto == LineSpacingTypeAuto.AutoAfter ) ? "afterAutospacing" : "beforeAutospacing";

        spacing.SetAttributeValue( XName.Get( spacingTypeAttribute, Document.w.NamespaceName ), spacingValue );
        spacing.SetAttributeValue( XName.Get( autoSpacingTypeAttribute, Document.w.NamespaceName ), 1 );

        if( spacingTypeAuto == LineSpacingTypeAuto.Auto )
        {
          spacing.SetAttributeValue( XName.Get( "after", Document.w.NamespaceName ), spacingValue );
          spacing.SetAttributeValue( XName.Get( "afterAutospacing", Document.w.NamespaceName ), 1 );
        }
      }
    }

    public Paragraph AppendBookmark( string bookmarkName )
    {
      XElement wBookmarkStart = new XElement(
          XName.Get( "bookmarkStart", Document.w.NamespaceName ),
          new XAttribute( XName.Get( "id", Document.w.NamespaceName ), bookmarkIdCounter ),
          new XAttribute( XName.Get( "name", Document.w.NamespaceName ), bookmarkName ) );
      Xml.Add( wBookmarkStart );

      XElement wBookmarkEnd = new XElement(
          XName.Get( "bookmarkEnd", Document.w.NamespaceName ),
          new XAttribute( XName.Get( "id", Document.w.NamespaceName ), bookmarkIdCounter ),
          new XAttribute( XName.Get( "name", Document.w.NamespaceName ), bookmarkName ) );
      Xml.Add( wBookmarkEnd );

      ++bookmarkIdCounter;

      return this;
    }

    public void ClearBookmarks()
    {
      var bookmarkStarts = this.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) );
      if( bookmarkStarts.Count() > 0 )
      {
        bookmarkStarts.Remove();
      }
      var bookmarkEnds = this.Xml.Descendants( XName.Get( "bookmarkEnd", Document.w.NamespaceName ) );
      if( bookmarkEnds.Count() > 0 )
      {
        bookmarkEnds.Remove();
      }
    }

    public bool ValidateBookmark( string bookmarkName )
    {
      return GetBookmarks().Any( b => b.Name.Equals( bookmarkName ) );
    }

    public IEnumerable<Bookmark> GetBookmarks()
    {
      var bookmarks = new List<Bookmark>();

      var bookmarksStartXml = this.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) );
      var bookmarksEndXml = this.Xml.Descendants( XName.Get( "bookmarkEnd", Document.w.NamespaceName ) );

      foreach( var bookmarkStartXml in bookmarksStartXml )
      {
        bookmarks.Add( new Bookmark()
        {
          Name = bookmarkStartXml.Attribute( XName.Get( "name", Document.w.NamespaceName ) ).Value,
          Id = bookmarkStartXml.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value,
          Paragraph = this
        } );
      }

      foreach( var bookmarkEndXml in bookmarksEndXml )
      {
        var id = bookmarkEndXml.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value;

        // bookmarkStart is not part of this Paragraph, find it in Document.
        if( bookmarks.Where( b => b.Id == id ).Count() == 0 )
        {
          var bookmarkStart = this.Document.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "id", Document.w.NamespaceName ) )?.Value == id )
                                  .LastOrDefault();

          if( ( bookmarks != null ) && ( bookmarkStart != null ) )
          {
            bookmarks.Add( new Bookmark()
            {
              Name = bookmarkStart.Attribute( XName.Get( "name", Document.w.NamespaceName ) ).Value,
              Id = id,
              Paragraph = this
            } );
          }
        }
      }

      return bookmarks;
    }

    public void InsertAtBookmark( string toInsert, string bookmarkName, Formatting formatting = null )
    {
      var bookmark = this.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) )
                          .Where( x => x.Attribute( XName.Get( "name", Document.w.NamespaceName ) ).Value == bookmarkName ).SingleOrDefault();
      var refPosition = bookmark;

      // bookmarkStart is not part of this paragraph, look for it in Document.
      if( bookmark == null )
      {
        var bookmarkStart = this.Document.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "name", Document.w.NamespaceName ) )?.Value == bookmarkName )
                                  .FirstOrDefault();
        if( bookmarkStart != null )
        {
          var bookmarkStartId = bookmarkStart.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value;

          bookmark = this.Xml.Descendants( XName.Get( "bookmarkEnd", Document.w.NamespaceName ) )
                          .Where( x => x.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value == bookmarkStartId ).SingleOrDefault();
          if( bookmark != null )
          {
            refPosition = this.Xml.Element( XName.Get( "r", Document.w.NamespaceName ) );
          }
        }
      }

      if( refPosition != null )
      {
        var run = HelperFunctions.FormatInput( toInsert, ( formatting != null ) ? formatting.Xml : null );
        refPosition.AddBeforeSelf( run );
        _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();
      }
    }

    public void ReplaceAtBookmark( string text, string bookmarkName, Formatting formatting = null )
    {
      string bookmarkStartId = null;
      XNode nextNode = null;

      var rList = new List<XElement>();
      var bookmarkStart = this.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "name", Document.w.NamespaceName ) ).Value == bookmarkName )
                                  .FirstOrDefault();
      if( bookmarkStart != null )
      {
        bookmarkStartId = bookmarkStart.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value;
        nextNode = bookmarkStart.NextNode;
      }
      // bookmarkStart is not in paragraph, look for bookmarkEnd.
      else
      {
        var bookmarkEnds = this.Xml.Descendants( XName.Get( "bookmarkEnd", Document.w.NamespaceName ) );
        if( bookmarkEnds.Count() > 1 )
          throw new InvalidDataException( "Unsupported exception: Paragraph do not contains the expected bookmarkStart and contains more than 1 bookmarkEnd." );
        if( bookmarkEnds.Count() == 0 )
          throw new InvalidDataException( "Unsupported exception: Paragraph do not contains a bookmarkStart or a bookmarkEnd." );

        bookmarkStartId = bookmarkEnds.First().Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value;
        nextNode = this.Xml.Element( XName.Get( "r", Document.w.NamespaceName ) );
        bookmarkStart = this.Xml.Element( XName.Get( "pPr", Document.w.NamespaceName ) );
      }

      XElement nextXElement = null;
      while( nextNode != null )
      {
        nextXElement = nextNode as XElement;
        if( ( nextXElement != null ) && ( nextXElement.Name.NamespaceName == Document.w.NamespaceName ) )
        {
          if( nextXElement.Name.LocalName == "r" )
          {
            rList.Add( nextXElement );
          }
          else if( ( nextXElement.Name.LocalName == "bookmarkEnd" )
                  && ( nextXElement.Attribute( XName.Get( "id", Document.w.NamespaceName ) )?.Value == bookmarkStartId ) )
            break;
        }

        nextNode = nextNode.NextNode;
      }

      if( nextXElement == null )
        return;

      if( rList.Count == 0 )
      {
        this.ReplaceAtBookmark_Core( text, bookmarkStart, formatting );
        return;
      }

      var tXElementFilled = false;
      foreach( var r in rList )
      {
        var tXElement = r.Elements( XName.Get( "t", Document.w.NamespaceName ) ).FirstOrDefault();
        if( tXElement == null )
        {
          if( !tXElementFilled )
          {
            this.ReplaceAtBookmark_Core( text, bookmarkStart, formatting );
            tXElementFilled = true;
          }
        }
        else
        {
          if( !tXElementFilled )
          {
            this.ReplaceAtBookmark_Core( text, r, formatting );
            if( formatting != null )
            {
              var rPr = r.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
              if( rPr != null )
              {
                rPr.Remove();
              }
            }
            tXElementFilled = true;
          }
          tXElement.Remove();
        }
      }
    }

    public void RemoveBookmark( string bookmarkName )
    {
      var bookmarkStartXml = this.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "name", Document.w.NamespaceName ) ).Value == bookmarkName )
                                  .FirstOrDefault();
      // bookmarkStart is not part of this paragraph, look for it in Document.
      if( bookmarkStartXml == null )
      {
        bookmarkStartXml = this.Document.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "name", Document.w.NamespaceName ) )?.Value == bookmarkName )
                                  .FirstOrDefault();
      }

      if( bookmarkStartXml == null )
        return;

      var bookmarkStartId = bookmarkStartXml.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value;

      var bookmarkEndXml = this.Xml.Descendants( XName.Get( "bookmarkEnd", Document.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "id", Document.w.NamespaceName ) )?.Value == bookmarkStartId )
                                  .FirstOrDefault();
      // bookmarkEnd is not part of this paragraph, look for it in Document.
      if( bookmarkEndXml == null )
      {
        bookmarkEndXml = this.Document.Xml.Descendants( XName.Get( "bookmarkEnd", Document.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "id", Document.w.NamespaceName ) )?.Value == bookmarkStartId )
                                  .FirstOrDefault();
      }

      if( bookmarkEndXml == null )
        return;

      bookmarkStartXml.Remove();
      bookmarkEndXml.Remove();
    }

    public Paragraph KeepWithNextParagraph( bool keepWithNextParagraph = true )
    {
      var pPr = GetOrCreate_pPr();
      var keepNextElement = pPr.Element( XName.Get( "keepNext", Document.w.NamespaceName ) );

      if( keepNextElement == null && keepWithNextParagraph )
      {
        pPr.Add( new XElement( XName.Get( "keepNext", Document.w.NamespaceName ) ) );
      }

      if( !keepWithNextParagraph && keepNextElement != null )
      {
        keepNextElement.Remove();
      }

      return this;
    }

    public Paragraph KeepLinesTogether( bool keepLinesTogether = true )
    {
      var pPr = GetOrCreate_pPr();
      var keepLinesElement = pPr.Element( XName.Get( "keepLines", Document.w.NamespaceName ) );

      if( keepLinesElement == null && keepLinesTogether )
      {
        pPr.Add( new XElement( XName.Get( "keepLines", Document.w.NamespaceName ) ) );
      }

      if( !keepLinesTogether )
      {
        keepLinesElement?.Remove();
      }

      return this;
    }

    [Obsolete( "Instead use : InsertHorizontalLine( HorizontalBorderPosition position, BorderStyle borderStyle, int size, int space, Color? color )" )]
    public void InsertHorizontalLine( HorizontalBorderPosition position = HorizontalBorderPosition.bottom, string lineType = "single", int size = 6, int space = 1, string color = "auto" )
    {
      var pBrXName = XName.Get( "pBdr", Document.w.NamespaceName );
      var borderPositionXName = ( position == HorizontalBorderPosition.bottom ) ? XName.Get( "bottom", Document.w.NamespaceName ) : XName.Get( "top", Document.w.NamespaceName );

      var pPr = this.GetOrCreate_pPr();
      var pBdr = pPr.Element( pBrXName );
      if( pBdr == null )
      {
        //Add border
        pPr.Add( new XElement( pBrXName ) );
        pBdr = pPr.Element( pBrXName );

        //Add bottom
        pBdr.Add( new XElement( borderPositionXName ) );
        var border = pBdr.Element( borderPositionXName );

        //Set border's attribute
        border.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), lineType );
        border.SetAttributeValue( XName.Get( "sz", Document.w.NamespaceName ), size.ToString() );
        border.SetAttributeValue( XName.Get( "space", Document.w.NamespaceName ), space.ToString() );
        border.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), color.Replace( "#", "" ) );
      }
    }

    public void InsertHorizontalLine( HorizontalBorderPosition position = HorizontalBorderPosition.bottom, BorderStyle lineType = BorderStyle.Tcbs_single, int size = 6, int space = 1, Xceed.Drawing.Color? color = null )
    {
      var pBrXName = XName.Get( "pBdr", Document.w.NamespaceName );
      var borderPositionXName = ( position == HorizontalBorderPosition.bottom ) ? XName.Get( "bottom", Document.w.NamespaceName ) : XName.Get( "top", Document.w.NamespaceName );

      var pPr = this.GetOrCreate_pPr();
      var pBdr = pPr.Element( pBrXName );
      if( pBdr == null )
      {
        //Add border
        pPr.Add( new XElement( pBrXName ) );
        pBdr = pPr.Element( pBrXName );

        //Add bottom
        pBdr.Add( new XElement( borderPositionXName ) );
        var border = pBdr.Element( borderPositionXName );

        //Set border's attribute
        border.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), lineType.ToString().Substring( 5 ) );
        border.SetAttributeValue( XName.Get( "sz", Document.w.NamespaceName ), size );
        border.SetAttributeValue( XName.Get( "space", Document.w.NamespaceName ), space );
        border.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), color.HasValue ? color.Value.ToHex() : "auto" );
      }
    }

    #endregion

    #region Internal Methods

    internal static void ResetDefaultValues()
    {
      Paragraph.DefaultLineSpacing = Paragraph.DefaultSingleLineSpacing;
      Paragraph.DefaultLineSpacingAfter = 0f;
      Paragraph.DefaultLineSpacingBefore = 0f;

      Paragraph.DefaultIndentationFirstLine = 0f;
      Paragraph.DefaultIndentationHanging = 0f;
      Paragraph.DefaultIndentationBefore = 0f;
      Paragraph.DefaultIndentationAfter = 0f;
    }

    internal static void SetDefaultValues( XElement pPr )
    {

      if( pPr == null )
        return;

      // Default line spacings.
      var spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );
      if( spacing != null )
      {
        var line = spacing.Attribute( XName.Get( "line", Document.w.NamespaceName ) );
        if( line != null )
        {
          float f;

          if( HelperFunctions.TryParseFloat( line.Value, out f ) )
          {
            Paragraph.DefaultLineSpacing = f / 20.0f;
          }
        }
        var after = spacing.Attribute( XName.Get( "after", Document.w.NamespaceName ) );
        if( after != null )
        {
          float f;
          if( HelperFunctions.TryParseFloat( after.Value, out f ) )
          {
            Paragraph.DefaultLineSpacingAfter = f / 20.0f;
          }
        }
        var before = spacing.Attribute( XName.Get( "before", Document.w.NamespaceName ) );
        if( before != null )
        {
          float f;
          if( HelperFunctions.TryParseFloat( before.Value, out f ) )
          {
            Paragraph.DefaultLineSpacingBefore = f / 20.0f;
          }
        }
        var lineRule = spacing.Attribute( XName.Get( "lineRule", Document.w.NamespaceName ) );
        if( lineRule != null )
        {
          Paragraph.DefaultLineRuleAuto = ( lineRule.Value == "auto" );
        }
      }

      // Default indentations.
      var ind = pPr.Element( XName.Get( "ind", Document.w.NamespaceName ) );
      if( ind != null )
      {
        var firstLine = ind.Attribute( XName.Get( "firstLine", Document.w.NamespaceName ) );
        if( firstLine != null )
        {
          Paragraph.DefaultIndentationFirstLine = float.Parse( firstLine.Value ) / 20f;
        }
        var hanging = ind.Attribute( XName.Get( "hanging", Document.w.NamespaceName ) );
        if( hanging != null )
        {
          Paragraph.DefaultIndentationHanging = float.Parse( hanging.Value ) / 20f;
        }
        var before = ind.Attribute( XName.Get( "left", Document.w.NamespaceName ) );
        if( before != null )
        {
          Paragraph.DefaultIndentationBefore = float.Parse( before.Value ) / 20f;
        }
        var after = ind.Attribute( XName.Get( "right", Document.w.NamespaceName ) );
        if( after != null )
        {
          Paragraph.DefaultIndentationAfter = float.Parse( after.Value ) / 20f;
        }
      }
    }

    internal void UpdateObjects( RemoveParagraphFlags removeParagraphFlags )
    {
      var previousParagraph = this.PreviousParagraph;

      if( this.FollowingTables != null )
      {
        foreach( var table in this.FollowingTables.ToList() )
        {
          if( ( removeParagraphFlags & RemoveParagraphFlags.Tables ) == RemoveParagraphFlags.Tables )
          {
            this.FollowingTables.Remove( table );

            table.Remove();
          }
          else if( previousParagraph != null )
          {
            if( previousParagraph.FollowingTables == null )
            {
              previousParagraph.FollowingTables = new List<Table>();
            }
            previousParagraph.FollowingTables.Add( table );
          }
        }
      }

      foreach( var picture in this.Pictures )
      {
        if( ( removeParagraphFlags & RemoveParagraphFlags.Pictures ) == RemoveParagraphFlags.Pictures )
        {
          picture.Remove();
        }
        else if( previousParagraph != null )
        {
          previousParagraph.Xml.Add( picture.Xml );
        }
      }

    }

    internal XElement GetOrCreate_pPr()
    {
      // Get the element.
      var pPr = Xml.Element( XName.Get( "pPr", Document.w.NamespaceName ) );

      // If it dosen't exist, create it.
      if( pPr == null )
      {
        Xml.AddFirst( new XElement( XName.Get( "pPr", Document.w.NamespaceName ) ) );
        pPr = Xml.Element( XName.Get( "pPr", Document.w.NamespaceName ) );
      }

      // Return the pPr element for this Paragraph.
      return pPr;
    }

    internal XElement GetOrCreate_rPr()
    {
      // Get the element.
      var rPr = Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );

      // If it dosen't exist, create it.
      if( rPr == null )
      {
        this.Xml.AddFirst( new XElement( XName.Get( "rPr", Document.w.NamespaceName ) ) );
        rPr = this.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
      }

      // Return the rPr element for this Paragraph.
      return rPr;
    }

    internal XElement GetOrCreate_pPr_ind()
    {
      // Get the element.
      XElement pPr = GetOrCreate_pPr();
      XElement ind = pPr.Element( XName.Get( "ind", Document.w.NamespaceName ) );

      // If it dosen't exist, create it.
      if( ind == null )
      {
        pPr.Add( new XElement( XName.Get( "ind", Document.w.NamespaceName ) ) );
        ind = pPr.Element( XName.Get( "ind", Document.w.NamespaceName ) );
      }

      // Return the pPr element for this Paragraph.
      return ind;
    }

    internal int GetTabStopPositionsCount()
    {
      var pPr = this.GetOrCreate_pPr();
      var tabs = pPr.Element( XName.Get( "tabs", Document.w.NamespaceName ) );
      if( tabs == null )
        return 0;

      return tabs.Elements().Count();
    }

    internal void RemoveHyperlinkRecursive( XElement xml, int index, ref int count, ref bool found )
    {
      if( xml.Name.LocalName.Equals( "hyperlink", StringComparison.CurrentCultureIgnoreCase ) )
      {
        // This is the hyperlink to be removed.
        if( count == index )
        {
          found = true;
          xml.Remove();
        }

        else
          count++;
      }

      if( xml.HasElements )
        foreach( XElement e in xml.Elements() )
          if( !found )
            RemoveHyperlinkRecursive( e, index, ref count, ref found );
    }

    internal void ResetBackers()
    {
      ParagraphNumberPropertiesBacker = null;
      IsListItemBacker = null;
      IndentLevelBacker = null;
    }

    static internal Picture CreatePicture( Document document, string id, string name, string descr, float width, float height )
    {
      var part = document._package.GetPart( document.PackagePart.GetRelationship( id ).TargetUri );
      long cx, cy;

      using( PackagePartStream packagePartStream = new PackagePartStream( part.GetStream() ) )
      {
        using( var img = Xceed.Drawing.Image.FromStream( packagePartStream ) )
        {
          // ooxml uses image size in EMU : 
          // image in inches(in) is : pt / 72
          // image in EMU is : in * 914400
          var imgHorizontalResolution = ( img.HorizontalResolution > 0 ) ? img.HorizontalResolution : Paragraph.DefaultImageHorizontalResolution;
          var imgVerticalResolution = ( img.VerticalResolution > 0 ) ? img.VerticalResolution : Paragraph.DefaultImageVerticalResolution;

          cx = Convert.ToInt64( img.Width * ( 72f / imgHorizontalResolution ) * Picture.EmusInPixel );
          cy = Convert.ToInt64( img.Height * ( 72f / imgVerticalResolution ) * Picture.EmusInPixel );
        }
      }

      var xml = XElement.Parse
        ( string.Format( @"
        <w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
            <w:drawing>
                <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                    <wp:extent cx=""{0}"" cy=""{1}"" />
                    <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0"" />
                    <wp:docPr id=""0"" name=""{3}"" descr=""{4}"" />
                    <wp:cNvGraphicFramePr>
                        <a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1"" />
                    </wp:cNvGraphicFramePr>
                    <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                        <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                            <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                <pic:nvPicPr>
                                  <pic:cNvPr id=""0"" name=""{3}"" />
                                <pic:cNvPicPr />
                                </pic:nvPicPr>
                                <pic:blipFill>
                                    <a:blip r:embed=""{2}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>
                                    <a:stretch>
                                        <a:fillRect />
                                    </a:stretch>
                                </pic:blipFill>
                                <pic:spPr>
                                    <a:xfrm>
                                        <a:off x=""0"" y=""0"" />
                                        <a:ext cx=""{0}"" cy=""{1}"" />
                                    </a:xfrm>
                                    <a:prstGeom prst=""rect"">
                                        <a:avLst />
                                    </a:prstGeom>
                                </pic:spPr>
                            </pic:pic>
                        </a:graphicData>
                    </a:graphic>
                </wp:inline>
            </w:drawing>
        </w:r>
        ", cx, cy, id, name, descr ) );

      var picture = new Picture( document, xml, new Image( document, document.PackagePart.GetRelationship( id ) ) );
      if( width > -1f )
      {
        picture.Width = width;
      }
      if( height > -1f )
      {
        picture.Height = height;
      }
      return picture;
    }

    internal Run GetFirstRunEffectedByEdit( int index, EditType type = EditType.ins )
    {
      int len = HelperFunctions.GetText( Xml ).Length;

      // Make sure we are looking within an acceptable index range.
      if( index < 0 || ( ( type == EditType.ins && index > len ) || ( type == EditType.del && index >= len ) ) )
        throw new ArgumentOutOfRangeException();

      // Need some memory that can be updated by the recursive search for the XElement to Split.
      int count = 0;
      Run theOne = null;

      GetFirstRunEffectedByEditRecursive( Xml, index, ref count, ref theOne, type );

      return theOne;
    }

    internal static bool CanReadXml( XElement xml )
    {
      return ( !xml.Name.Equals( XName.Get( "drawing", Document.w.NamespaceName ) ) && !xml.Name.Equals( XName.Get( "Fallback", Document.mc.NamespaceName ) ) );
    }

    internal void GetFirstRunEffectedByEditRecursive( XElement Xml, int index, ref int count, ref Run theOne, EditType type )
    {
      if( CanReadXml( Xml ) )
      {
        count += HelperFunctions.GetSize( Xml );
      }
      else
      {
        return;
      }

      // If the EditType is deletion then we must return the next blah
      if( count > 0 && ( ( type == EditType.del && count > index ) || ( type == EditType.ins && count >= index ) ) )
      {
        // Correct the index
        foreach( XElement e in Xml.ElementsBeforeSelf() )
        {
          count -= HelperFunctions.GetSize( e );
        }

        count -= HelperFunctions.GetSize( Xml );
        count = Math.Max( 0, count );

        // We have found the element, now find the run it belongs to.
        while( ( Xml.Name.LocalName != "r" ) )
        {
          Xml = Xml.Parent;
          if( Xml == null )
            return;
        }

        theOne = new Run( Document, Xml, count );
        return;
      }

      // Ignore Fallback and drawing to be symmetric with HelperFunctions.GetTextRecursive.
      var fallbackValue = Xml.Name.Equals(XName.Get("Fallback", Document.mc.NamespaceName));
      var drawingValue = Xml.Name.Equals(XName.Get("drawing", Document.w.NamespaceName));

      if( Xml.HasElements && !fallbackValue && !drawingValue )
      {
        foreach( XElement e in Xml.Elements() )
        {
          if( theOne == null )
          {
            this.GetFirstRunEffectedByEditRecursive( e, index, ref count, ref theOne, type );
          }
        }
      }
    }

    static internal int GetElementTextLength( XElement xml )
    {
      int count = 0;

      if( xml == null )
        return count;

      // Increment count for empty paragraphs as well
      if( xml.Name.LocalName == "p" && xml.Descendants( XName.Get( "t", Document.w.NamespaceName ) ).Count() == 0 )
      {
        count++;
      }
      else
      {
        foreach( var d in xml.Descendants() )
        {
          switch( d.Name.LocalName )
          {
            case "tab":
              if( d.Parent.Name.LocalName != "tabs" )
                count++;
              break;
            case "br":
              // Manage only line Breaks.
              if( HelperFunctions.IsLineBreak( d ) )
                count++;
              break;
            case "t":
              goto case "delText";
            case "delText":
              count += d.Value.Length;
              break;
            default:
              break;
          }
        }
      }

      return count;
    }

    internal XElement[] SplitEdit( XElement edit, int index, EditType type )
    {
      Run run = GetFirstRunEffectedByEdit( index, type );

      XElement[] splitRun = Run.SplitRun( run, index, type );

      XElement splitLeft = new XElement( edit.Name, edit.Attributes(), run.Xml.ElementsBeforeSelf(), splitRun[ 0 ] );
      if( GetElementTextLength( splitLeft ) == 0 )
        splitLeft = null;

      XElement splitRight = new XElement( edit.Name, edit.Attributes(), splitRun[ 1 ], run.Xml.ElementsAfterSelf() );
      if( GetElementTextLength( splitRight ) == 0 )
        splitRight = null;

      return
      (
          new XElement[]
          {
                    splitLeft,
                    splitRight
          }
      );
    }

    internal void ApplyTextFormattingProperty( XName textFormatPropName, string value, object content )
    {
      XElement rPr = null;

      if( _runs.Count == 0 )
      {
        var pPr = this.Xml.Element( XName.Get( "pPr", Document.w.NamespaceName ) );
        if( pPr == null )
        {
          this.Xml.AddFirst( new XElement( XName.Get( "pPr", Document.w.NamespaceName ) ) );
          pPr = this.Xml.Element( XName.Get( "pPr", Document.w.NamespaceName ) );
        }

        rPr = pPr.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
        if( rPr == null )
        {
          pPr.AddFirst( new XElement( XName.Get( "rPr", Document.w.NamespaceName ) ) );
          rPr = pPr.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
        }

        rPr.SetElementValue( textFormatPropName, value );

        var lastElement = rPr.Elements( textFormatPropName ).Last();
        // Check if the content is an attribute
        if( content as XAttribute != null )
        {
          // Add or Update the attribute to the last element
          if( lastElement.Attribute( ( (XAttribute)( content ) ).Name ) == null )
          {
            lastElement.Add( content );
          }
          else
          {
            lastElement.Attribute( ( (XAttribute)( content ) ).Name ).Value = ( (XAttribute)( content ) ).Value;
          }
        }
        return;
      }

      var isFontPropertiesList = false;
      var fontProperties = content as IEnumerable;
      if( fontProperties != null )
      {
        foreach( object property in fontProperties )
        {
          isFontPropertiesList = ( property as XAttribute != null );
        }
      }

      foreach( XElement run in _runs )
      {
        rPr = run.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
        if( rPr == null )
        {
          run.AddFirst( new XElement( XName.Get( "rPr", Document.w.NamespaceName ) ) );
          rPr = run.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
        }

        rPr.SetElementValue( textFormatPropName, value );
        var last = rPr.Elements( textFormatPropName ).Last();

        if( isFontPropertiesList )
        {
          foreach( object property in fontProperties )
          {
            if( last.Attribute( ( (XAttribute)( property ) ).Name ) == null )
            {
              last.Add( property );
            }
            else
            {
              last.Attribute( ( (XAttribute)( property ) ).Name ).Value = ( (XAttribute)( property ) ).Value;
            }
          }
        }


        if( content as XAttribute != null )//If content is an attribute
        {
          if( last.Attribute( ( (XAttribute)( content ) ).Name ) == null )
          {
            last.Add( content ); //Add this attribute if element doesn't have it
          }
          else
          {
            last.Attribute( ( (XAttribute)( content ) ).Name ).Value = ( (XAttribute)( content ) ).Value; //Apply value only if element already has it
          }
        }
        else
        {
          //IMPORTANT
          //But what to do if it is not?
        }
      }
    }

    internal bool IsLineSpacingRuleAuto()
    {
      var pPr = GetOrCreate_pPr();
      var spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

      if( spacing != null )
      {
        var lineRule = spacing.Attribute( XName.Get( "lineRule", Document.w.NamespaceName ) );
        if( lineRule != null )
        {
          return ( lineRule.Value == "auto" );  //"auto" for for Single, double, 1.5, multiple. Not "auto" for "exact"/"atLeast".
        }
      }

      return Paragraph.DefaultLineRuleAuto;
    }

    internal bool IsLineSpacingRuleExactlyOrAtLeast()
    {
      var pPr = GetOrCreate_pPr();
      var spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

      if( spacing != null )
      {
        var lineRule = spacing.Attribute( XName.Get( "lineRule", Document.w.NamespaceName ) );
        if( lineRule != null )
        {
          return ( ( lineRule.Value == "exact" ) || ( lineRule.Value == "atLeast" ) );
        }
      }

      return false;
    }

    internal bool IsLineSpacingRuleExactly()
    {
      var pPr = GetOrCreate_pPr();
      var spacing = pPr.Element( XName.Get( "spacing", Document.w.NamespaceName ) );

      if( spacing != null )
      {
        var lineRule = spacing.Attribute( XName.Get( "lineRule", Document.w.NamespaceName ) );
        if( lineRule != null )
        {
          return ( lineRule.Value == "exact" );
        }
      }

      return false;
    }

    internal void ClearMagicTextCache()
    {
      _magicText = null;
    }

    internal bool CanAddAttribute( XAttribute att )
    {
      if( ( att.Name.LocalName == "hanging" ) && ( this.IndentationFirstLine != Paragraph.DefaultIndentationFirstLine ) )
        return false;

      return true;
    }

    internal int GetNumId()
    {
      var numIdNode = this.Xml.Descendants().FirstOrDefault( s => s.Name.LocalName == "numId" );
      if( numIdNode == null )
        return -1;
      return Int32.Parse( numIdNode.Attribute( Document.w + "val" ).Value );
    }

    internal int GetListItemLevel()
    {
      if( this.ParagraphNumberProperties != null )
      {
        var ilvl = this.ParagraphNumberProperties.Descendants().FirstOrDefault( el => el.Name.LocalName == "ilvl" );
        if( ilvl != null )
        {
          return Int32.Parse( ilvl.Attribute( Document.w + "val" ).Value );
        }
      }
      return -1;
    }

    internal void SetAsBookmark( string bookmarkName )
    {
      XElement wBookmarkStart = new XElement(
          XName.Get( "bookmarkStart", Document.w.NamespaceName ),
          new XAttribute( XName.Get( "id", Document.w.NamespaceName ), bookmarkIdCounter ),
          new XAttribute( XName.Get( "name", Document.w.NamespaceName ), bookmarkName ) );
      Xml.AddFirst( wBookmarkStart );  //Add as First Element

      XElement wBookmarkEnd = new XElement(
          XName.Get( "bookmarkEnd", Document.w.NamespaceName ),
          new XAttribute( XName.Get( "id", Document.w.NamespaceName ), bookmarkIdCounter ),
          new XAttribute( XName.Get( "name", Document.w.NamespaceName ), bookmarkName ) );
      Xml.Add( wBookmarkEnd );  //Add as last element

      ++bookmarkIdCounter;
    }

    internal bool IsInTOC()
    {
      var sdtContent = this.Xml.Parent;
      if( ( sdtContent != null ) && ( sdtContent.Name == XName.Get( "sdtContent", Document.w.NamespaceName ) ) )
      {
        var sdt = sdtContent.Parent;
        if( ( sdt != null ) && ( sdt.Name == XName.Get( "sdt", Document.w.NamespaceName ) ) )
        {
          return ( sdt.Descendants( XName.Get( "docPartGallery", Document.w.NamespaceName ) )
                      .FirstOrDefault( x => x.Attribute( XName.Get( "val", Document.w.NamespaceName ) ).Value == "Table of Contents" ) != null );
        }
      }
      return false;
    }

    internal bool IsInTOCVisible()
    {
      if( this.IsInTOC() )
      {
        var sdtContent = this.Xml.Parent;
        if( ( sdtContent != null ) && ( sdtContent.Name == XName.Get( "sdtContent", Document.w.NamespaceName ) ) )
        {
          if( this.StyleId.StartsWith( "TOC" ) )
          {
            var styleDigit = this.StyleId.Where( c => char.IsDigit( c ) );
            // TOCHeading
            if( styleDigit.Count() <= 0 )
              return true;

            var sdt = sdtContent.Parent;
            if( ( sdt != null ) && ( sdt.Name == XName.Get( "sdt", Document.w.NamespaceName ) ) )
            {
              var tocSwitches = sdt.Descendants( XName.Get( "instrText", Document.w.NamespaceName ) ).FirstOrDefault( instrText => instrText.Value.Contains( "TOC" ) );
              if( tocSwitches != null )
              {
                var includedHeadingStylesIndex = tocSwitches.Value.IndexOf( "o" );
                if( includedHeadingStylesIndex >= 0 )
                {
                  var bound1 = tocSwitches.Value.IndexOf( "\"", includedHeadingStylesIndex ) + 1;
                  var bound2 = tocSwitches.Value.IndexOf( "\"", bound1 );
                  var includedHeadingStyles = tocSwitches.Value.Substring( bound1, bound2 - bound1 );

                  var maxStyle = int.Parse( includedHeadingStyles[ includedHeadingStyles.Length - 1 ].ToString() );
                  int counter = 1;
                  var styleDigitValue = int.Parse( string.Concat( styleDigit ) );
                  while( counter <= maxStyle )
                  {
                    if( styleDigitValue == counter )
                      return true;
                    ++counter;
                  }

                  return false;
                }
              }
            }
          }
        }
      }

      return true;
    }

    internal List<XElement> GetSdtContentRuns()
    {
      List<XElement> result = null;

      var sdtContents = this.Xml.Descendants( XName.Get( "sdtContent", Document.w.NamespaceName ) );
      if( sdtContents != null )
      {
        foreach( var sdtContent in sdtContents )
        {
          var runs = sdtContent.Elements( XName.Get( "r", Document.w.NamespaceName ) );
          if( runs != null && runs.Count() > 0 )
          {
            if( result == null )
            {
              result = runs.ToList();
            }
            else
            {
              result.AddRange( runs );
            }
          }
        }
      }

      return result;
    }

    internal bool IsInSdt()
    {
      return ( this.GetParentSdt() != null );
    }

    internal XElement GetParentSdt()
    {
      return this.Xml.Ancestors( XName.Get( "sdt", Document.w.NamespaceName ) ).FirstOrDefault();
    }

    #endregion

    #region Private Methods

    private void ApplyFormattingFrom( ref Formatting newFormatting, Formatting sourceFormatting )
    {
      //Set the formatting properties of clone based on received formatting.
      newFormatting.FontFamily = sourceFormatting.FontFamily;
      newFormatting.Language = sourceFormatting.Language;
      if( sourceFormatting.Bold.HasValue )
      {
        newFormatting.Bold = sourceFormatting.Bold;
      }
      if( sourceFormatting.CapsStyle.HasValue )
      {
        newFormatting.CapsStyle = sourceFormatting.CapsStyle;
      }
      if( sourceFormatting.FontColor.HasValue )
      {
        newFormatting.FontColor = sourceFormatting.FontColor;
      }
      if( sourceFormatting.Hidden.HasValue )
      {
        newFormatting.Hidden = sourceFormatting.Hidden;
      }
      if( sourceFormatting.Highlight.HasValue )
      {
        newFormatting.Highlight = sourceFormatting.Highlight;
      }
      if( sourceFormatting.Italic.HasValue )
      {
        newFormatting.Italic = sourceFormatting.Italic;
      }
      if( sourceFormatting.Kerning.HasValue )
      {
        newFormatting.Kerning = sourceFormatting.Kerning;
      }
      if( sourceFormatting.Misc.HasValue )
      {
        newFormatting.Misc = sourceFormatting.Misc;
      }
      if( sourceFormatting.PercentageScale.HasValue )
      {
        newFormatting.PercentageScale = sourceFormatting.PercentageScale;
      }
      if( sourceFormatting.Position.HasValue )
      {
        newFormatting.Position = sourceFormatting.Position;
      }
      if( sourceFormatting.Script.HasValue )
      {
        newFormatting.Script = sourceFormatting.Script;
      }
      if( sourceFormatting.Size.HasValue )
      {
        newFormatting.Size = sourceFormatting.Size;
      }
      if( sourceFormatting.Spacing.HasValue )
      {
        newFormatting.Spacing = sourceFormatting.Spacing;
      }
      if( sourceFormatting.StrikeThrough.HasValue )
      {
        newFormatting.StrikeThrough = sourceFormatting.StrikeThrough;
      }
      if( sourceFormatting.UnderlineColor.HasValue )
      {
        newFormatting.UnderlineColor = sourceFormatting.UnderlineColor;
      }
      if( sourceFormatting.UnderlineStyle.HasValue )
      {
        newFormatting.UnderlineStyle = sourceFormatting.UnderlineStyle;
      }
    }

    private void RebuildDocProperties()
    {
      if( this.Xml != null )
      {
        docProperties =
        (
            from xml in Xml.Descendants( XName.Get( "fldSimple", Document.w.NamespaceName ) )
            select new DocProperty( Document, xml )
        ).ToList();
      }
    }

    private XElement GetParagraphNumberProperties()
    {
      var numPrNode = this.Xml.Descendants().FirstOrDefault( el => el.Name.LocalName == "numPr" );
      if( numPrNode != null )
      {
        // numId of 0 is not a ListItem.
        if( this.GetNumId() == 0 )
          return null;
      }
      else
      {
        // Look in style and basedOn styles of this paragraph.
        var paragraphStyle = HelperFunctions.GetParagraphStyleFromStyleId( this.Document, this.StyleId );
        numPrNode = this.GetParagraphNumberPropertiesFromStyle( paragraphStyle );
      }

      return numPrNode;
    }

    private XElement GetParagraphNumberPropertiesFromStyle( XElement style )
    {
      if( style == null )
        return null;

      var numPrNode = style.Descendants().FirstOrDefault( el => el.Name.LocalName == "numPr" );
      if( numPrNode != null )
        return numPrNode;

      var basedOn = style.Element( XName.Get( "basedOn", Document.w.NamespaceName ) );
      if( basedOn != null )
      {
        var val = basedOn.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
        if( val != null )
        {
          var basedOnParagraphStyle = HelperFunctions.GetParagraphStyleFromStyleId( this.Document, val.Value );
          return this.GetParagraphNumberPropertiesFromStyle( basedOnParagraphStyle );
        }
      }

      return null;
    }

    private List<Picture> GetPictures( string localName, string localNameEquals, string attributeName )
    {
      // Do not include picture contained in a Fallback.
      var pictures =
       (
           from p in Xml.Descendants()
           where ( p.Name.LocalName == localName )
           where ( p.Ancestors().FirstOrDefault( x => x.Name.Equals( XName.Get( "Fallback", Document.mc.NamespaceName ) ) ) == null )
           let ids =
           (
               from e in p.Descendants()
               where e.Name.LocalName.Equals( localNameEquals )
               select e.Attribute( XName.Get( attributeName, "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) ).Value
           )
           where ( ids != null ) && ( ids.Count() > 0 )
           from id in ids
           where ( this.PackagePart.RelationshipExists( id ) )
           select new Picture( this.Document, p, new Image( this.Document, this.PackagePart.GetRelationship( id ) ) ) { PackagePart = this.PackagePart }

       ).ToList();

      return pictures;
    }

    private IList<Paragraph> GetHeaderParagraphs( Paragraph paragraph )
    {
      var header = this.Document.Headers.First;

      if( header != null )
      {
        if( header.PackagePart.Uri == paragraph.PackagePart.Uri )
          return header.Paragraphs;
      }

      header = this.Document.Headers.Odd;
      if( header != null )
      {
        if( header.PackagePart.Uri == paragraph.PackagePart.Uri )
          return header.Paragraphs;
      }

      header = this.Document.Headers.Even;
      if( header != null )
        return header.Paragraphs;

      return null;
    }

    private IList<Paragraph> GetFooterParagraphs( Paragraph paragraph )
    {
      var footer = this.Document.Footers.First;

      if( footer != null )
      {
        if( footer.PackagePart.Uri == paragraph.PackagePart.Uri )
          return footer.Paragraphs;
      }

      footer = this.Document.Footers.Odd;
      if( footer != null )
      {
        if( footer.PackagePart.Uri == paragraph.PackagePart.Uri )
          return footer.Paragraphs;
      }

      footer = this.Document.Footers.Even;
      if( footer != null )
        return footer.Paragraphs;

      return null;
    }

    private void ClearContainerParagraphsCache()
    {
      switch( this.ParentContainer )
      {
        case ContainerType.Header:
          {
            this.ClearHeaderParagraphsCache();
            break;
          }
          ;

        case ContainerType.Footer:
          {
            this.ClearFooterParagraphsCache();
            break;
          }

        case ContainerType.Body:
          {
            this.Document.ClearParagraphsCache();
            break;
          }

        default:
          {
            // A paragraph can be located inside a shape/table, ...
            // A shape/table can be located inside a header/footer/body
            var parentContainers = this.Xml.Ancestors();

            if( parentContainers.FirstOrDefault( parent => parent.Name.LocalName == "hdr" ) != null )
            {
              this.ClearHeaderParagraphsCache();
            }
            else if( parentContainers.FirstOrDefault( parent => parent.Name.LocalName == "ftr" ) != null )
            {
              this.ClearFooterParagraphsCache();
            }
            else
            {
              this.Document.ClearParagraphsCache();
            }

            break;
          }
      }
    }

    private void ClearHeaderParagraphsCache()
    {
      var header = this.Document.Headers.First;

      if( header != null )
      {
        if( header.PackagePart.Uri == this.PackagePart.Uri )
        {
          header.ClearParagraphsCache();
          return;
        }
      }

      header = this.Document.Headers.Odd;
      if( header != null )
      {
        if( header.PackagePart.Uri == this.PackagePart.Uri )
        {
          header.ClearParagraphsCache();
          return;
        }
      }

      header = this.Document.Headers.Even;
      header.ClearParagraphsCache();
    }

    private void ClearFooterParagraphsCache()
    {
      var footer = this.Document.Footers.First;

      if( footer != null )
      {
        if( footer.PackagePart.Uri == this.PackagePart.Uri )
        {
          footer.ClearParagraphsCache();
          return;
        }
      }

      footer = this.Document.Footers.Odd;
      if( footer != null )
      {
        if( footer.PackagePart.Uri == this.PackagePart.Uri )
        {
          footer.ClearParagraphsCache();
          return;
        }
      }

      footer = this.Document.Footers.Even;
      footer.ClearParagraphsCache();
    }

    private void ValidateInsert()
    {
      if( m_removed )
      {
        throw new InvalidOperationException( "Cannot insert before or after a removed paragraph." );
      }
    }


























    private void ReplaceAtBookmark_Core( string text, XElement bookmark, Formatting formatting = null )
    {
      var xElementList = HelperFunctions.FormatInput( text, ( formatting != null ) ? formatting.Xml : null );
      bookmark.AddAfterSelf( xElementList );

      _runs = this.Xml.Descendants( XName.Get( "r", Document.w.NamespaceName ) ).ToList();
    }

    private void AddParagraphStyleIfNotPresent( string wantedParagraphStyleId )
    {
      if( string.IsNullOrEmpty( wantedParagraphStyleId ) )
        return;

      // Load _styles if not loaded.
      if( this.Document._styles == null )
      {
        PackagePart word_styles = this.Document._package.GetPart( new Uri( "/word/styles.xml", UriKind.Relative ) );
        using( TextReader tr = new StreamReader( word_styles.GetStream() ) )
        {
          this.Document._styles = XDocument.Load( tr );
        }
      }

      // Check if this Paragraph StyleName exists in _styles.
      var paragraphStyleExist = HelperFunctions.GetParagraphStyleFromStyleId( this.Document, wantedParagraphStyleId ) != null;

      // This Paragraph StyleId doesn't exists in _styles, add it.
      if( !paragraphStyleExist )
      {
        // Load the default_styles.
        var stylesDoc = HelperFunctions.DecompressXMLResource( HelperFunctions.GetResources( ResourceType.DefaultStyle ) );

        // get the paragraph styles.
        var availableParagraphStyles =
         (
             from s in stylesDoc.Element( Document.w + "styles" ).Elements( Document.w + "style" )
             let type = s.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
             where ( type != null && type.Value == "paragraph" )
             select s
         );

        // Get the wanted Paragraph style.
        var wantedParagraphStyle =
         (
             from s in availableParagraphStyles
             let styleId = s.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) )
             where ( styleId != null && styleId.Value == wantedParagraphStyleId )
             select s
         );

        // Add the wanted paragraph style to _styles.
        if( ( wantedParagraphStyle != null ) && ( wantedParagraphStyle.Count() > 0 ) )
        {
          this.Document._styles.Element( Document.w + "styles" ).Add( wantedParagraphStyle );
        }
      }
    }

    private XElement GetNumberContentBasedOnLast_rPr()
    {
      var rPr = this.Xml.Descendants( XName.Get( "rPr", Document.w.NamespaceName ) ).LastOrDefault();
      var rPrText = ( rPr != null ) ? rPr.ToString() : "<w:rPr><w:noProof/></w:rPr>";

      var content = XElement.Parse( string.Format( @"
              <w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   {0}
                   <w:t>1</w:t> 
               </w:r>"
        , rPrText )
      );

      return content;
    }

    private bool UpdateTextToReplace( ReplaceTextOptionsBase replaceTextOptions, ref string textToAnalyse )
    {
      if( ( replaceTextOptions.StartIndex >= 0 ) && ( replaceTextOptions.EndIndex >= 0 ) && ( replaceTextOptions.StartIndex >= replaceTextOptions.EndIndex ) )
        throw new InvalidDataException( "replaceTextOptions.EndIndex must be greater than replaceTextOptions.StartIndex." );

      if( replaceTextOptions.StartIndex >= 0 )
      {
        if( this.EndIndex < replaceTextOptions.StartIndex )
          return false;

        if( this.StartIndex < replaceTextOptions.StartIndex )
        {
          var start = Math.Max( this.StartIndex, replaceTextOptions.StartIndex );
          var end = ( replaceTextOptions.EndIndex > 0 ) ? Math.Min( this.EndIndex, replaceTextOptions.EndIndex ) : this.EndIndex;

          textToAnalyse = textToAnalyse.Substring( start - this.StartIndex, end - this.StartIndex - ( start - this.StartIndex ) );

          if( replaceTextOptions.EndIndex > 0 )
            return true;
        }
      }

      if( replaceTextOptions.EndIndex >= 0 )
      {
        if( this.StartIndex > replaceTextOptions.EndIndex )
          return false;

        if( this.EndIndex > replaceTextOptions.EndIndex )
        {
          var start = ( replaceTextOptions.StartIndex > 0 ) ? Math.Max( this.StartIndex, replaceTextOptions.StartIndex ) : this.StartIndex;
          var end = Math.Min( this.EndIndex, replaceTextOptions.EndIndex );

          textToAnalyse = textToAnalyse.Substring( start - this.StartIndex, end - this.StartIndex - ( start - this.StartIndex ) );
        }
      }

      return true;
    }

    private bool ReplaceTextCore( Match singleMatch, StringReplaceTextOptions replaceTextOptions )
    {
      // Assume the formatting matches until proven otherwise.
      bool formattingMatch = true;
      var singleMatchIndexInFullText = ( replaceTextOptions.StartIndex < 0 )
                                      ? singleMatch.Index
                                      : ( this.StartIndex < replaceTextOptions.StartIndex ) ? singleMatch.Index + Math.Abs( this.StartIndex - replaceTextOptions.StartIndex ) : singleMatch.Index;

      // Does the user want to match formatting?
      if( replaceTextOptions.FormattingToMatch != null )
      {
        // The number of characters processed so far
        int processed = 0;

        do
        {
          // Get the next run effected
          var run = GetFirstRunEffectedByEdit( singleMatchIndexInFullText + processed );

          // Get this runs properties
          var rPr = run.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );

          if( rPr == null )
          {
            rPr = new Formatting().Xml;
          }

          /* 
           * Make sure that every formatting element in f.xml is also in this run,
           * if this is not true, then their formatting does not match.
           */
          if( !HelperFunctions.ContainsEveryChildOf( replaceTextOptions.FormattingToMatch.Xml, rPr, replaceTextOptions.FormattingToMatchOptions ) )
          {
            formattingMatch = false;
            break;
          }

          // We have processed some characters, so update the counter.
          processed += run.Value.Length;

        } while( processed < singleMatch.Length );
      }

      // If the formatting matches, do the replace.
      if( formattingMatch )
      {
        //perform RegEx substitutions. Only named groups are not supported. Everything else is supported. However character escapes are not covered.
        if( replaceTextOptions.UseRegExSubstitutions && !string.IsNullOrEmpty( replaceTextOptions.NewValue ) )
        {
          replaceTextOptions.NewValue = replaceTextOptions.NewValue.Replace( "$&", singleMatch.Value );
          if( singleMatch.Groups.Count > 0 )
          {
            int lastcap = 0;
            for( int k = 0; k < singleMatch.Groups.Count; k++ )
            {
              var g = singleMatch.Groups[ k ];
              if( ( g == null ) || ( g.Value == "" ) )
                continue;
              replaceTextOptions.NewValue = replaceTextOptions.NewValue.Replace( "$" + k.ToString(), g.Value );
              lastcap = k;
            }
            replaceTextOptions.NewValue = replaceTextOptions.NewValue.Replace( "$+", singleMatch.Groups[ lastcap ].Value );
          }
          if( singleMatchIndexInFullText > 0 )
          {
            replaceTextOptions.NewValue = replaceTextOptions.NewValue.Replace( "$`", this.Text.Substring( 0, singleMatchIndexInFullText ) );
          }
          if( ( singleMatchIndexInFullText + singleMatch.Length ) < this.Text.Length )
          {
            replaceTextOptions.NewValue = replaceTextOptions.NewValue.Replace( "$'", this.Text.Substring( singleMatchIndexInFullText + singleMatch.Length ) );
          }
          replaceTextOptions.NewValue = replaceTextOptions.NewValue.Replace( "$_", this.Text );
          replaceTextOptions.NewValue = replaceTextOptions.NewValue.Replace( "$$", "$" );
        }

        var replacedSuccess = false;
        if( !string.IsNullOrEmpty( replaceTextOptions.NewValue ) )
        {
          this.InsertText( singleMatchIndexInFullText + singleMatch.Length, replaceTextOptions.NewValue, replaceTextOptions.TrackChanges, replaceTextOptions.NewFormatting );
          replacedSuccess = true;
        }
        if( singleMatch.Length > 0 )
        {
          this.RemoveText( singleMatchIndexInFullText, singleMatch.Length, replaceTextOptions.TrackChanges, replaceTextOptions.RemoveEmptyParagraph );
          replacedSuccess = true;
        }

        return replacedSuccess;
      }

      return false;
    }

    private bool ReplaceTextCore( Match singleMatch, FunctionReplaceTextOptions replaceTextOptions )
    {
      var formattingMatch = true;
      var singleMatchIndexInFullText = ( replaceTextOptions.StartIndex < 0 )
                                     ? singleMatch.Index
                                     : ( this.StartIndex < replaceTextOptions.StartIndex ) ? singleMatch.Index + Math.Abs( this.StartIndex - replaceTextOptions.StartIndex ) : singleMatch.Index;

      if( replaceTextOptions.FormattingToMatch != null )
      {
        int processed = 0;

        while( processed < singleMatch.Length )
        {
          var run = this.GetFirstRunEffectedByEdit( singleMatchIndexInFullText + processed );
          var rPr = run.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
          if( rPr == null )
          {
            rPr = new Formatting().Xml;
          }

          // Make sure that every formatting element in matchFormatting.Xml is also in this run,
          // if false => formatting does not match.
          if( !HelperFunctions.ContainsEveryChildOf( replaceTextOptions.FormattingToMatch.Xml, rPr, replaceTextOptions.FormattingToMatchOptions ) )
          {
            formattingMatch = false;
            break;
          }

          processed += run.Value.Length;
        }
      }

      // Replace text when formatting matches.
      if( formattingMatch )
      {
        int lastcap = 0;
        for( int k = 0; k < singleMatch.Groups.Count; k++ )
        {
          var g = singleMatch.Groups[ k ];
          if( ( g == null ) || ( g.Value == "" ) )
            continue;
          lastcap = k;
        }

        var replacedSuccess = false;
        var newValue = replaceTextOptions.RegexMatchHandler.Invoke( singleMatch.Groups[ lastcap ].Value );
        if( !string.IsNullOrEmpty( newValue ) )
        {
          this.InsertText( singleMatchIndexInFullText + singleMatch.Value.Length, newValue, replaceTextOptions.TrackChanges, replaceTextOptions.NewFormatting );
          replacedSuccess = true;
        }
        if( singleMatch.Length > 0 )
        {
          this.RemoveText( singleMatchIndexInFullText, singleMatch.Value.Length, replaceTextOptions.TrackChanges, replaceTextOptions.RemoveEmptyParagraph );
          replacedSuccess = true;
        }

        return replacedSuccess;
      }

      return false;
    }

    private bool ReplaceTextCore( Match singleMatch, ObjectReplaceTextOptions replaceTextOptions )
    {
      // Assume the formatting matches until proven otherwise.
      bool formattingMatch = true;
      var singleMatchIndexInFullText = ( replaceTextOptions.StartIndex < 0 )
                                     ? singleMatch.Index
                                     : ( this.StartIndex < replaceTextOptions.StartIndex ) ? singleMatch.Index + Math.Abs( this.StartIndex - replaceTextOptions.StartIndex ) : singleMatch.Index;

      // Does the user want to match formatting?
      if( replaceTextOptions.FormattingToMatch != null )
      {
        // The number of characters processed so far
        int processed = 0;

        do
        {
          // Get the next run effected
          var run = GetFirstRunEffectedByEdit( singleMatchIndexInFullText + processed );

          // Get this runs properties
          var rPr = run.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) );

          if( rPr == null )
          {
            rPr = new Formatting().Xml;
          }

          /* 
           * Make sure that every formatting element in f.xml is also in this run,
           * if this is not true, then their formatting does not match.
           */
          if( !HelperFunctions.ContainsEveryChildOf( replaceTextOptions.FormattingToMatch.Xml, rPr, replaceTextOptions.FormattingToMatchOptions ) )
          {
            formattingMatch = false;
            break;
          }

          // We have processed some characters, so update the counter.
          processed += run.Value.Length;

        } while( processed < singleMatch.Length );
      }

      // If the formatting matches, do the replace.
      if( formattingMatch )
      {
        var replacedSuccess = false;

        if( replaceTextOptions.NewObject != null )
        {
          if( replaceTextOptions.NewObject is Picture )
          {
            this.InsertPicture( (Picture)replaceTextOptions.NewObject, singleMatchIndexInFullText + singleMatch.Length );
            replacedSuccess = true;
          }
          else if( replaceTextOptions.NewObject is Hyperlink )
          {
            this.InsertHyperlink( (Hyperlink)replaceTextOptions.NewObject, singleMatchIndexInFullText + singleMatch.Length );
            replacedSuccess = true;
          }
          else if( replaceTextOptions.NewObject is Table )
          {
            this.InsertTableAfterSelf( (Table)replaceTextOptions.NewObject );
            replacedSuccess = true;
          }
          else
          {
            throw new ArgumentException( "Unknown object received. Valid objects are Picture, Hyperlink or Table." );
          }
        }
        if( singleMatch.Length > 0 )
        {
          this.RemoveText( singleMatchIndexInFullText, singleMatch.Length, replaceTextOptions.TrackChanges, replaceTextOptions.RemoveEmptyParagraph );
          replacedSuccess = true;
        }

        return replacedSuccess;
      }

      return false;
    }



































    private bool IsInMainContainer()
    {
      return ( ( this.ParentContainer == ContainerType.Body )
            || ( this.ParentContainer == ContainerType.Header )
            || ( this.ParentContainer == ContainerType.Footer ) );
    }

    private XElement SetPageCountFields( PageNumberFormat? pnf = null, bool useSectionPageCount = false )
    {
      var fldSimple = new XElement( XName.Get( "fldSimple", Document.w.NamespaceName ) );
      var fields = useSectionPageCount ? @" SECTIONPAGES " : @" NUMPAGES ";

      if( pnf == PageNumberFormat.normal )
      {
        fldSimple.Add( new XAttribute( XName.Get( "instr", Document.w.NamespaceName ), fields + @"  \* MERGEFORMAT " ) );
      }
      else if( pnf == PageNumberFormat.roman )
      {
        fldSimple.Add( new XAttribute( XName.Get( "instr", Document.w.NamespaceName ), fields + @" \* ROMAN  \* MERGEFORMAT " ) );
      }
      else
      {
        fldSimple.Add( new XAttribute( XName.Get( "instr", Document.w.NamespaceName ), fields ) );
      }

      return fldSimple;
    }

    private XElement SetPageNumberFields( PageNumberFormat? pnf = null )
    {
      var fldSimple = new XElement( XName.Get( "fldSimple", Document.w.NamespaceName ) );

      if( pnf == PageNumberFormat.normal )
      {
        fldSimple.Add( new XAttribute( XName.Get( "instr", Document.w.NamespaceName ), @" PAGE \* MERGEFORMAT " ) );
      }
      else if( pnf == PageNumberFormat.roman )
      {
        fldSimple.Add( new XAttribute( XName.Get( "instr", Document.w.NamespaceName ), @" PAGE  \* ROMAN  \* MERGEFORMAT " ) );
      }
      else
      {
        fldSimple.Add( new XAttribute( XName.Get( "instr", Document.w.NamespaceName ), @" PAGE " ) );
      }

      return fldSimple;
    }

    internal void NeedRefreshIndexes()
    {
      var mainParentContainer = this.GetMainParentContainer();
      if( mainParentContainer != null )
      {
        mainParentContainer.NeedRefreshParagraphIndexes = true;
      }
    }

    internal Container GetMainParentContainer()
    {
      if( this.ParentContainer == ContainerType.Header )
        return this.GetHeaderContainer();
      else if( this.ParentContainer == ContainerType.Footer )
        return this.GetFooterContainer();
      else if( this.ParentContainer == ContainerType.Body )
        return this.Document;

      // A paragraph can be located inside a shape/table, ...
      // A shape/table can be located inside a header/footer/body
      var parentContainers = this.Xml.Ancestors();

      if( parentContainers.FirstOrDefault( parent => parent.Name.LocalName == "hdr" ) != null )
        return this.GetHeaderContainer();
      else if( parentContainers.FirstOrDefault( parent => parent.Name.LocalName == "ftr" ) != null )
        return this.GetFooterContainer();

      return this.Document;
    }

    private Container GetHeaderContainer()
    {
      foreach( var section in this.Document.Sections )
      {
        var header = section.Headers.First;
        if( header != null )
        {
          if( header.PackagePart.Uri == this.PackagePart.Uri )
            return header;
        }

        header = section.Headers.Odd;
        if( header != null )
        {
          if( header.PackagePart.Uri == this.PackagePart.Uri )
          {
            return header;
          }
        }

        header = section.Headers.Even;
        return header;
      }

      return null;
    }

    private Container GetFooterContainer()
    {
      foreach( var section in this.Document.Sections )
      {
        var footer = section.Footers.First;
        if( footer != null )
        {
          if( footer.PackagePart.Uri == this.PackagePart.Uri )
            return footer;
        }

        footer = section.Footers.Odd;
        if( footer != null )
        {
          if( footer.PackagePart.Uri == this.PackagePart.Uri )
          {
            return footer;
          }
        }

        footer = section.Footers.Even;
        return footer;
      }
      return null;
    }

    private IList<Paragraph> GetContainerParagraphs()
    {
      if( this.ParentContainer == ContainerType.Header )
      {
        return this.GetHeaderParagraphs( this );
      }
      else if( this.ParentContainer == ContainerType.Footer )
      {
        return this.GetFooterParagraphs( this );
      }
      else if( this.ParentContainer == ContainerType.Body )
      {
        return this.Document.Paragraphs;
      }
      else
      {
        // A paragraph can be located inside a shape/table, ...
        // A shape/table can be located inside a header/footer/body
        var parentContainers = this.Xml.Ancestors();

        if( parentContainers.FirstOrDefault( parent => parent.Name.LocalName == "hdr" ) != null )
        {
          return this.GetHeaderParagraphs( this );
        }
        else if( parentContainers.FirstOrDefault( parent => parent.Name.LocalName == "ftr" ) != null )
        {
          return this.GetFooterParagraphs( this );
        }
        else
        {
          return this.Document.Paragraphs;
        }
      }
    }

    private void InsertFollowingTables( Paragraph p, bool isInsertingAfter )
    {
      if( p != null )
      {
        var tables = p.FollowingTables;

        if( tables != null )
        {
          if( isInsertingAfter )
          {
            foreach( var table in tables )
            {
              this.InsertTableAfterSelf( table );
            }
          }
          else
          {
            foreach( var table in tables )
            {
              this.InsertTableBeforeSelf( table );
            }
          }
        }
      }
    }

    #endregion
  }

  public class Run : DocumentElement
  {
    #region Private Members

    // A lookup for the text elements in this paragraph
    private Dictionary<int, Text> textLookup = new Dictionary<int, Text>();

    private int startIndex;
    private int endIndex;
    private string text;

    #endregion

    #region Public Properties

    public int StartIndex
    {
      get
      {
        return startIndex;
      }
    }

    public int EndIndex
    {
      get
      {
        return endIndex;
      }
    }

    #endregion

    #region Internal Properties

    internal string Value
    {
      set
      {
        text = value;
      }
      get
      {
        return text;
      }
    }

    #endregion

    #region Constructors

    internal Run( Document document, XElement xml, int startIndex )
        : base( document, xml )
    {
      this.startIndex = startIndex;

      // Get the text elements in this run
      IEnumerable<XElement> texts = xml.Descendants();

      int start = startIndex;

      // Loop through each text in this run
      foreach( XElement te in texts )
      {
        switch( te.Name.LocalName )
        {
          case "tab":
            {
              textLookup.Add( start + 1, new Text( Document, te, start ) );
              text += "\t";
              start++;
              break;
            }
          case "br":
            {
              // Manage only line Breaks.
              if( HelperFunctions.IsLineBreak( te ) )
              {
                textLookup.Add( start + 1, new Text( Document, te, start ) );
                text += "\n";
                start++;
              }

              break;
            }
          case "t":
            goto case "delText";
          case "delText":
            {
              // Only add strings which are not empty
              if( te.Value.Length > 0 )
              {
                textLookup.Add( start + te.Value.Length, new Text( Document, te, start ) );
                text += te.Value;
                start += te.Value.Length;
              }
              break;
            }
          default:
            break;
        }
      }

      endIndex = start;
    }

    #endregion

    #region Internal Methods

    static internal XElement[] SplitRun( Run r, int index, EditType type = EditType.ins )
    {
      index = index - r.StartIndex;

      Text t = r.GetFirstTextEffectedByEdit( index, type );
      XElement[] splitText = Text.SplitText( t, index );

      XElement splitLeft = new XElement( r.Xml.Name, r.Xml.Attributes(), r.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) ), t.Xml.ElementsBeforeSelf().Where( n => n.Name.LocalName != "rPr" ), splitText[ 0 ] );
      if( Paragraph.GetElementTextLength( splitLeft ) == 0 )
        splitLeft = null;

      XElement splitRight = new XElement( r.Xml.Name, r.Xml.Attributes(), r.Xml.Element( XName.Get( "rPr", Document.w.NamespaceName ) ), splitText[ 1 ], t.Xml.ElementsAfterSelf().Where( n => n.Name.LocalName != "rPr" ) );
      if( Paragraph.GetElementTextLength( splitRight ) == 0 )
        splitRight = null;

      return
      (
          new XElement[]
          {
                    splitLeft,
                    splitRight
          }
      );
    }

    internal Text GetFirstTextEffectedByEdit( int index, EditType type = EditType.ins )
    {
      // Make sure we are looking within an acceptable index range.
      if( index < 0 || index > HelperFunctions.GetText( Xml ).Length )
        throw new ArgumentOutOfRangeException();

      // Need some memory that can be updated by the recursive search for the XElement to Split.
      int count = 0;
      Text theOne = null;

      GetFirstTextEffectedByEditRecursive( Xml, index, ref count, ref theOne, type );

      return theOne;
    }

    internal void GetFirstTextEffectedByEditRecursive( XElement Xml, int index, ref int count, ref Text theOne, EditType type = EditType.ins )
    {
      count += HelperFunctions.GetSize( Xml );
      if( count > 0 && ( ( type == EditType.del && count > index ) || ( type == EditType.ins && count >= index ) ) )
      {
        theOne = new Text( Document, Xml, count - HelperFunctions.GetSize( Xml ) );
        return;
      }

      if( Xml.HasElements )
        foreach( XElement e in Xml.Elements() )
          if( theOne == null )
            GetFirstTextEffectedByEditRecursive( e, index, ref count, ref theOne );
    }

    #endregion
  }

  internal class Text : DocumentElement
  {
    #region Private Members

    private int startIndex;
    private int endIndex;
    private string text;

    #endregion

    #region Public Properties

    public int StartIndex
    {
      get
      {
        return startIndex;
      }
    }

    public int EndIndex
    {
      get
      {
        return endIndex;
      }
    }

    public string Value
    {
      get
      {
        return text;
      }
    }

    #endregion

    #region Constructors

    internal Text( Document document, XElement xml, int startIndex )
        : base( document, xml )
    {
      this.startIndex = startIndex;

      switch( Xml.Name.LocalName )
      {
        case "t":
          {
            goto case "delText";
          }

        case "delText":
          {
            endIndex = startIndex + xml.Value.Length;
            text = xml.Value;
            break;
          }

        case "br":
          {
            // Manage only line Breaks.
            if( HelperFunctions.IsLineBreak( Xml ) )
            {
              text = "\n";
              endIndex = startIndex + 1;
            }
            break;
          }

        case "tab":
          {
            text = "\t";
            endIndex = startIndex + 1;
            break;
          }
        default:
          {
            break;
          }
      }
    }

    #endregion

    #region Public Methods

    public static void PreserveSpace( XElement e )
    {
      // PreserveSpace should only be used on (t or delText) elements
      if( !e.Name.Equals( Document.w + "t" ) && !e.Name.Equals( Document.w + "delText" ) )
        throw new ArgumentException( "SplitText can only split elements of type t or delText", "e" );

      // Check if this w:t contains a space atribute
      XAttribute space = e.Attributes().Where( a => a.Name.Equals( XNamespace.Xml + "space" ) ).SingleOrDefault();

      // This w:t's text begins or ends with whitespace
      if( e.Value.StartsWith( " " ) || e.Value.EndsWith( " " ) )
      {
        // If this w:t contains no space attribute, add one.
        if( space == null )
          e.Add( new XAttribute( XNamespace.Xml + "space", "preserve" ) );
      }

      // This w:t's text does not begin or end with a space
      else
      {
        // If this w:r contains a space attribute, remove it.
        if( space != null )
          space.Remove();
      }
    }

    #endregion

    #region Internal Methods

    internal static XElement[] SplitText( Text t, int index )
    {
      if( index < t.startIndex || index > t.EndIndex )
        throw new ArgumentOutOfRangeException( "index" );

      XElement splitLeft = null;
      XElement splitRight = null;
      if( t.Xml.Name.LocalName == "t" || t.Xml.Name.LocalName == "delText" )
      {
        // The origional text element, now containing only the text before the index point.
        splitLeft = new XElement( t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring( 0, index - t.startIndex ) );
        if( splitLeft.Value.Length == 0 )
        {
          splitLeft = null;
        }
        else
        {
          Text.PreserveSpace( splitLeft );
        }

        // The origional text element, now containing only the text after the index point.
        splitRight = new XElement( t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring( index - t.startIndex, t.Xml.Value.Length - ( index - t.startIndex ) ) );
        if( splitRight.Value.Length == 0 )
        {
          splitRight = null;
        }
        else
        {
          Text.PreserveSpace( splitRight );
        }
      }

      else
      {
        if( index == t.EndIndex )
        {
          splitLeft = t.Xml;
        }
        else
        {
          splitRight = t.Xml;
        }
      }

      return
      (
          new XElement[]
          {
                    splitLeft,
                    splitRight
          }
      );
    }

    #endregion
  }
}
