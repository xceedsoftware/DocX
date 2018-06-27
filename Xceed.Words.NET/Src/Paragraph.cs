/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Security.Principal;
using System.IO.Packaging;
using System.Drawing;
using System.Globalization;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Represents a document paragraph.
  /// </summary>
  public class Paragraph : InsertBeforeOrAfter
  {

    #region Internal Members

    // The Append family of functions use this List to apply style.
    internal List<XElement> _runs;
    internal int _startIndex, _endIndex;
    internal List<XElement> _styles = new List<XElement>();

    internal const float DefaultLineSpacing = 1.1f * 20.0f;

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
    private Table followingTable;

    #endregion

    #region Private Properties
    private XElement ParagraphNumberPropertiesBacker
    {
      get; set;
    }

    private bool? IsListItemBacker
    {
      get; set;
    }

    private int? IndentLevelBacker
    {
      get; set;
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

    /// <summary>
    /// Returns a list of all Pictures in a Paragraph.
    /// </summary>
    /// <example>
    /// Returns a list of all Pictures in a Paragraph.
    /// <code>
    /// <![CDATA[
    /// // Create a document.
    /// using (DocX document = DocX.Load(@"Test.docx"))
    /// {
    ///    // Get the first Paragraph in a document.
    ///    Paragraph p = document.Paragraphs[0];
    /// 
    ///    // Get all of the Pictures in this Paragraph.
    ///    List<Picture> pictures = p.Pictures;
    ///
    ///    // Save this document.
    ///    document.Save();
    /// }
    /// ]]>
    /// </code>
    /// </example>
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

    /// <summary>
    /// Returns a list of Hyperlinks in this Paragraph.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Load(@"Test.docx"))
    /// {
    ///    // Get the first Paragraph in this document.
    ///    Paragraph p = document.Paragraphs[0];
    ///    
    ///    // Get all of the hyperlinks in this Paragraph.
    ///    <![CDATA[ List<hyperlink> ]]> hyperlinks = paragraph.Hyperlinks;
    ///    
    ///    // Change the first hyperlinks text and Uri
    ///    Hyperlink h0 = hyperlinks[0];
    ///    h0.Text = "DocX";
    ///    h0.Uri = new Uri("http://docx.codeplex.com");
    ///
    ///    // Save this document.
    ///    document.Save();
    /// }
    /// </code>
    /// </example>
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
            foreach( XElement r in e.ElementsAfterSelf( XName.Get( "r", DocX.w.NamespaceName ) ) )
            {
              // Add this run to the list.
              hyperlink_runs.Add( r );

              var fldChar = r.Descendants( XName.Get( "fldChar", DocX.w.NamespaceName ) ).SingleOrDefault<XElement>();
              if( fldChar != null )
              {
                var fldCharType = fldChar.Attribute( XName.Get( "fldCharType", DocX.w.NamespaceName ) );
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

    ///<summary>
    /// The style name of the paragraph.
    ///</summary>
    public string StyleName
    {
      get
      {
        var element = this.GetOrCreate_pPr();
        var styleElement = element.Element( XName.Get( "pStyle", DocX.w.NamespaceName ) );
        var attr = styleElement?.Attribute( XName.Get( "val", DocX.w.NamespaceName ) );
        if ( !string.IsNullOrEmpty( attr?.Value ) )
        {
          return attr.Value;
        }
        return "Normal";
      }
      set
      {
        if( string.IsNullOrEmpty( value ) )
        {
          value = "Normal";
        }
        var element = this.GetOrCreate_pPr();
        var styleElement = element.Element( XName.Get( "pStyle", DocX.w.NamespaceName ) );
        if( styleElement == null )
        {
          element.Add( new XElement( XName.Get( "pStyle", DocX.w.NamespaceName ) ) );
          styleElement = element.Element( XName.Get( "pStyle", DocX.w.NamespaceName ) );
        }
        styleElement.SetAttributeValue( XName.Get( "val", DocX.w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Returns a list of field type DocProperty in this document.
    /// </summary>
    public List<DocProperty> DocumentProperties
    {
      get
      {
        return docProperties;
      }
    }

    /// <summary>
    /// Gets or Sets the Direction of content in this Paragraph.
    /// <example>
    /// Create a Paragraph with content that flows right to left. Default is left to right.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///     // Create a new Paragraph with the text "Hello World".
    ///     Paragraph p = document.InsertParagraph("Hello World.");
    /// 
    ///     // Make this Paragraph flow right to left. Default is left to right.
    ///     p.Direction = Direction.RightToLeft;
    ///     
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// </summary>
    public Direction Direction
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();
        XElement bidi = pPr.Element( XName.Get( "bidi", DocX.w.NamespaceName ) );
        return bidi == null ? Direction.LeftToRight : Direction.RightToLeft;
      }

      set
      {
        direction = value;

        XElement pPr = GetOrCreate_pPr();
        XElement bidi = pPr.Element( XName.Get( "bidi", DocX.w.NamespaceName ) );

        if( direction == Direction.RightToLeft )
        {
          if( bidi == null )
            pPr.Add( new XElement( XName.Get( "bidi", DocX.w.NamespaceName ) ) );
        }

        else
        {
          bidi?.Remove();
        }
      }
    }

    /// <summary>
    /// Get or set the indentation of the first line of this Paragraph.
    /// </summary>
    /// <example>
    /// Indent only the first line of a Paragraph.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///     // Create a new Paragraph.
    ///     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
    /// 
    ///     // Indent only the first line of the Paragraph.
    ///     p.IndentationFirstLine = 2.0f;
    ///     
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public float IndentationFirstLine
    {
      get
      {
        GetOrCreate_pPr();
        XElement ind = GetOrCreate_pPr_ind();
        XAttribute firstLine = ind.Attribute( XName.Get( "firstLine", DocX.w.NamespaceName ) );

        if( firstLine != null )
          return float.Parse( firstLine.Value ) / 570f;        

        return 0.0f;
      }

      set
      {
        if( IndentationFirstLine != value )
        {
          indentationFirstLine = value;

          GetOrCreate_pPr();
          XElement ind = GetOrCreate_pPr_ind();

          // Paragraph can either be firstLine or hanging (Remove hanging).
          XAttribute hanging = ind.Attribute( XName.Get( "hanging", DocX.w.NamespaceName ) );
          hanging?.Remove();

          string indentation = ( ( indentationFirstLine / 0.1 ) * 57 ).ToString();
          XAttribute firstLine = ind.Attribute( XName.Get( "firstLine", DocX.w.NamespaceName ) );
          if( firstLine != null )
            firstLine.Value = indentation;
          else
            ind.Add( new XAttribute( XName.Get( "firstLine", DocX.w.NamespaceName ), indentation ) );
        }
      }
    }

    /// <summary>
    /// Get or set the indentation of all but the first line of this Paragraph.
    /// </summary>
    /// <example>
    /// Indent all but the first line of a Paragraph.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///     // Create a new Paragraph.
    ///     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
    /// 
    ///     // Indent all but the first line of the Paragraph.
    ///     p.IndentationHanging = 1.0f;
    ///     
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public float IndentationHanging
    {
      get
      {
        GetOrCreate_pPr();
        var ind = GetOrCreate_pPr_ind();
        var hanging = ind.Attribute( XName.Get( "hanging", DocX.w.NamespaceName ) );

        if( hanging != null )
          return float.Parse( hanging.Value ) / 570f;

        return 0.0f;
      }

      set
      {
        if( IndentationHanging != value )
        {
          indentationHanging = value;

          GetOrCreate_pPr();
          var ind = GetOrCreate_pPr_ind();

          // Paragraph can either be firstLine or hanging (Remove firstLine).
          var firstLine = ind.Attribute( XName.Get( "firstLine", DocX.w.NamespaceName ) );
          if( firstLine != null )
          {
            firstLine.Remove();
          }

          var indentationValue = ( ( indentationHanging / 0.1 ) * 57 );
          var indentation = indentationValue.ToString();
          var hanging = ind.Attribute( XName.Get( "hanging", DocX.w.NamespaceName ) );
          if( hanging != null )
          {
            hanging.Value = indentation;
          }
          else
          {
            ind.Add( new XAttribute( XName.Get( "hanging", DocX.w.NamespaceName ), indentation ) );
          }
          IndentationBefore = indentationHanging;
        }
      }
    }

    /// <summary>
    /// Set the before indentation in cm for this Paragraph.
    /// </summary>
    /// <example>
    /// // Indent an entire Paragraph from the left.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///    // Create a new Paragraph.
    ///    Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
    ///
    ///    // Indent this entire Paragraph from the left.
    ///    p.IndentationBefore = 2.0f;
    ///    
    ///    // Save all changes made to this document.
    ///    document.Save();
    ///}
    /// </code>
    /// </example>
    public float IndentationBefore
    {
      get
      {
        GetOrCreate_pPr();
        var ind = GetOrCreate_pPr_ind();

        var left = ind.Attribute( XName.Get( "left", DocX.w.NamespaceName ) );
        if( left != null )
          return float.Parse( left.Value ) / 570f;

        return 0.0f;
      }

      set
      {
        if( IndentationBefore != value )
        {
          indentationBefore = value;

          GetOrCreate_pPr();
          var ind = GetOrCreate_pPr_ind();

          var indentation = ( ( indentationBefore / 0.1 ) * 57 ).ToString();

          var left = ind.Attribute( XName.Get( "left", DocX.w.NamespaceName ) );
          if( left != null )
          {
            left.Value = indentation;
          }
          else
          {
            ind.Add( new XAttribute( XName.Get( "left", DocX.w.NamespaceName ), indentation ) );
          }
        }
      }
    }


    /// <summary>
    /// Set the after indentation in cm for this Paragraph.
    /// </summary>
    /// <example>
    /// // Indent an entire Paragraph from the right.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///     // Create a new Paragraph.
    ///     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
    /// 
    ///     // Make the content of this Paragraph flow right to left.
    ///     p.Direction = Direction.RightToLeft;
    /// 
    ///     // Indent this entire Paragraph from the right.
    ///     p.IndentationAfter = 2.0f;
    ///     
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public float IndentationAfter
    {
      get
      {
        GetOrCreate_pPr();
        var ind = GetOrCreate_pPr_ind();

        var right = ind.Attribute( XName.Get( "right", DocX.w.NamespaceName ) );
        if( right != null )
          return float.Parse( right.Value ) / 570f;

        return 0.0f;
      }

      set
      {
        if( IndentationAfter != value )
        {
          indentationAfter = value;

          GetOrCreate_pPr();
          var ind = GetOrCreate_pPr_ind();

          var indentation = ( ( indentationAfter / 0.1 ) * 57 ).ToString();

          var right = ind.Attribute( XName.Get( "right", DocX.w.NamespaceName ) );
          if( right != null )
          {
            right.Value = indentation;
          }
          else
          {
            ind.Add( new XAttribute( XName.Get( "right", DocX.w.NamespaceName ), indentation ) );
          }
        }
      }
    }

    /// <summary>
    /// Gets or set this Paragraphs text alignment.
    /// </summary>
    public Alignment Alignment
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();
        XElement jc = pPr.Element( XName.Get( "jc", DocX.w.NamespaceName ) );

        if( jc != null )
        {
          XAttribute a = jc.Attribute( XName.Get( "val", DocX.w.NamespaceName ) );

          switch( a.Value.ToLower() )
          {
            case "left":
              return Xceed.Words.NET.Alignment.left;
            case "right":
              return Xceed.Words.NET.Alignment.right;
            case "center":
              return Xceed.Words.NET.Alignment.center;
            case "both":
              return Xceed.Words.NET.Alignment.both;
          }
        }

        return Xceed.Words.NET.Alignment.left;
      }

      set
      {
        alignment = value;

        XElement pPr = GetOrCreate_pPr();
        XElement jc = pPr.Element( XName.Get( "jc", DocX.w.NamespaceName ) );

        if( alignment != Xceed.Words.NET.Alignment.left )
        {
          if( jc == null )
            pPr.Add( new XElement( XName.Get( "jc", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), alignment.ToString() ) ) );
          else
            jc.Attribute( XName.Get( "val", DocX.w.NamespaceName ) ).Value = alignment.ToString();
        }

        else
        {
          if( jc != null )
            jc.Remove();
        }
      }
    }

    /// <summary>
    /// Gets the text value of this Paragraph.
    /// </summary>
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

    /// <summary>
    /// Gets the formatted text value of this Paragraph.
    /// </summary>
    public List<FormattedText> MagicText
    {
      // Returns the underlying XElement's Value property.
      get
      {
        return HelperFunctions.GetFormattedText( Xml );
      }
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <returns>This Paragraph in curent culture</returns>
    /// <example>
    /// Add a new Paragraph with russian text to this document and then set language of text to local culture.
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph with russian text and set curent local culture to it.
    ///     Paragraph p = document.InsertParagraph("Привет мир!").CurentCulture();
    ///       
    ///     // Save this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public Paragraph CurentCulture()
    {
      ApplyTextFormattingProperty( XName.Get( "lang", DocX.w.NamespaceName ),
          string.Empty,
          new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), CultureInfo.CurrentCulture.Name ) );
      return this;
    }

    ///<summary>
    /// Returns table following the paragraph. Null if the following element isn't table.
    ///</summary>
    public Table FollowingTable
    {
      get
      {
        return followingTable;
      }
      internal set
      {
        followingTable = value;
      }
    }

    public float LineSpacing
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();
        XElement spacing = pPr.Element( XName.Get( "spacing", DocX.w.NamespaceName ) );

        if( spacing != null )
        { 
          XAttribute line = spacing.Attribute( XName.Get( "line", DocX.w.NamespaceName ) );
          if( line != null )
          {
            float f;

            if( float.TryParse( line.Value, out f ) )
              return f / 20.0f;
          }
        }

        return Paragraph.DefaultLineSpacing;
      }

      set
      {
        Spacing( value );
      }
    }

    public float LineSpacingBefore
    {
      get
      {
        XElement pPr = GetOrCreate_pPr();
        XElement spacing = pPr.Element( XName.Get( "spacing", DocX.w.NamespaceName ) );

        if( spacing != null )
        {
          XAttribute line = spacing.Attribute( XName.Get( "before", DocX.w.NamespaceName ) );
          if( line != null )
          {
            float f;

            if( float.TryParse( line.Value, out f ) )
              return f / 20.0f;
          }
        }

        return 0.0f;
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
        XElement spacing = pPr.Element( XName.Get( "spacing", DocX.w.NamespaceName ) );

        if( spacing != null )
        {
          XAttribute line = spacing.Attribute( XName.Get( "after", DocX.w.NamespaceName ) );
          if( line != null )
          {
            float f;

            if( float.TryParse( line.Value, out f ) )
              return f / 20.0f;
          }
        }

        return 0.0f;
      }


      set
      {
        SpacingAfter( value );
      }
    }

    public XElement ParagraphNumberProperties
    {
      get
      {
        return ParagraphNumberPropertiesBacker ?? ( ParagraphNumberPropertiesBacker = GetParagraphNumberProperties() );
      }
    }

    /// <summary>
    /// Indicates if this paragraph is a list element
    /// </summary>
    public bool IsListItem
    {
      get
      {
        IsListItemBacker = IsListItemBacker ?? ( ParagraphNumberProperties != null );
        return ( bool )IsListItemBacker;
      }
    }

    /// <summary>
    /// Get the indentation level of the list item
    /// </summary>
    public int? IndentLevel
    {
      get
      {
        if( !IsListItem )
          return null;

        if( IndentLevelBacker != null )
          return IndentLevelBacker;

        var ilvl = ParagraphNumberProperties.Descendants().FirstOrDefault( el => el.Name.LocalName == "ilvl" );
        return IndentLevelBacker = ( ilvl != null ) ? int.Parse( ilvl.GetAttribute( DocX.w + "val" ) ) : 0;
      }
    }

    public bool IsKeepWithNext
    {
      get
      {
        var pPr = this.GetOrCreate_pPr();
        var keepNext = pPr.Element( XName.Get( "keepNext", DocX.w.NamespaceName ) );

        return ( keepNext != null );
      }
    }


    #endregion

    #region Constructors

    internal Paragraph( DocX document, XElement xml, int startIndex, ContainerType parentContainerType = ContainerType.None ) : base( document, xml )
    {
      _startIndex = startIndex;
      _endIndex = startIndex + GetElementTextLength( xml );

      ParentContainer = parentContainerType;

      RebuildDocProperties();

      //// Check if this Paragraph references any pStyle elements.
      //var stylesElements = xml.Descendants( XName.Get( "pStyle", DocX.w.NamespaceName ) );

      //// If one or more pStyles are referenced.
      //if( stylesElements.Count() > 0 )
      //{
      //  Uri style_package_uri = new Uri( "/word/styles.xml", UriKind.Relative );
      //  PackagePart styles_document = document.package.GetPart( style_package_uri );

      //  using( TextReader tr = new StreamReader( styles_document.GetStream() ) )
      //  {
      //    XDocument style_document = XDocument.Load( tr );
      //    XElement styles_element = style_document.Element( XName.Get( "styles", DocX.w.NamespaceName ) );

      //    var styles_element_ids = stylesElements.Select( e => e.Attribute( XName.Get( "val", DocX.w.NamespaceName ) ).Value );

      //    //foreach(string id in styles_element_ids)
      //    //{
      //    //    var style = 
      //    //    (
      //    //        from d in styles_element.Descendants()
      //    //        let styleId = d.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
      //    //        let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
      //    //        where type != null && type.Value == "paragraph" && styleId != null && styleId.Value == id
      //    //        select d
      //    //    ).First();

      //    //    styles.Add(style);
      //    //} 
      //  }
      //}

      _runs = Xml.Elements( XName.Get( "r", DocX.w.NamespaceName ) ).ToList();
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Insert a new Table before this Paragraph, this Table can be from this document or another document.
    /// </summary>
    /// <param name="t">The Table t to be inserted.</param>
    /// <returns>A new Table inserted before this Paragraph.</returns>
    /// <example>
    /// Insert a new Table before this Paragraph.
    /// <code>
    /// // Place holder for a Table.
    /// Table t;
    ///
    /// // Load document a.
    /// using (DocX documentA = DocX.Load(@"a.docx"))
    /// {
    ///     // Get the first Table from this document.
    ///     t = documentA.Tables[0];
    /// }
    ///
    /// // Load document b.
    /// using (DocX documentB = DocX.Load(@"b.docx"))
    /// {
    ///     // Get the first Paragraph in document b.
    ///     Paragraph p2 = documentB.Paragraphs[0];
    ///
    ///     // Insert the Table from document a before this Paragraph.
    ///     Table newTable = p2.InsertTableBeforeSelf(t);
    ///
    ///     // Save all changes made to document b.
    ///     documentB.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public override Table InsertTableBeforeSelf( Table t )
    {
      t = base.InsertTableBeforeSelf( t );
      t.PackagePart = this.PackagePart;
      return t;
    }

    /// <summary>
    /// Insert a new Table into this document before this Paragraph.
    /// </summary>
    /// <param name="rowCount">The number of rows this Table should have.</param>
    /// <param name="columnCount">The number of columns this Table should have.</param>
    /// <returns>A new Table inserted before this Paragraph.</returns>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     //Insert a Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("Hello World", false);
    ///
    ///     // Insert a new Table before this Paragraph.
    ///     Table newTable = p.InsertTableBeforeSelf(2, 2);
    ///     newTable.Design = TableDesign.LightShadingAccent2;
    ///     newTable.Alignment = Alignment.center;
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public override Table InsertTableBeforeSelf( int rowCount, int columnCount )
    {
      return base.InsertTableBeforeSelf( rowCount, columnCount );
    }

    /// <summary>
    /// Insert a new Table after this Paragraph.
    /// </summary>
    /// <param name="t">The Table t to be inserted.</param>
    /// <returns>A new Table inserted after this Paragraph.</returns>
    /// <example>
    /// Insert a new Table after this Paragraph.
    /// <code>
    /// // Place holder for a Table.
    /// Table t;
    ///
    /// // Load document a.
    /// using (DocX documentA = DocX.Load(@"a.docx"))
    /// {
    ///     // Get the first Table from this document.
    ///     t = documentA.Tables[0];
    /// }
    ///
    /// // Load document b.
    /// using (DocX documentB = DocX.Load(@"b.docx"))
    /// {
    ///     // Get the first Paragraph in document b.
    ///     Paragraph p2 = documentB.Paragraphs[0];
    ///
    ///     // Insert the Table from document a after this Paragraph.
    ///     Table newTable = p2.InsertTableAfterSelf(t);
    ///
    ///     // Save all changes made to document b.
    ///     documentB.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public override Table InsertTableAfterSelf( Table t )
    {
      t = base.InsertTableAfterSelf( t );
      t.PackagePart = this.PackagePart;
      return t;
    }

    /// <summary>
    /// Insert a new Table into this document after this Paragraph.
    /// </summary>
    /// <param name="rowCount">The number of rows this Table should have.</param>
    /// <param name="columnCount">The number of columns this Table should have.</param>
    /// <returns>A new Table inserted after this Paragraph.</returns>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     //Insert a Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("Hello World", false);
    ///
    ///     // Insert a new Table after this Paragraph.
    ///     Table newTable = p.InsertTableAfterSelf(2, 2);
    ///     newTable.Design = TableDesign.LightShadingAccent2;
    ///     newTable.Alignment = Alignment.center;
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public override Table InsertTableAfterSelf( int rowCount, int columnCount )
    {
      return base.InsertTableAfterSelf( rowCount, columnCount );
    }

    /// <summary>
    /// Replaces an existing Picture with a new Picture.
    /// </summary>
    /// <param name="toBeReplaced">The picture object to be replaced.</param>
    /// <param name="replaceWith">The picture object that should be inserted instead of <paramref name="toBeReplaced"/>.</param>
    /// <returns>The new <see cref="Picture"/> object that replaces the old one.</returns>
    public Picture ReplacePicture( Picture toBeReplaced, Picture replaceWith )
    {
      var document = this.Document;
      var newDocPrId = document.GetNextFreeDocPrId();

      var xml = XElement.Parse( toBeReplaced.Xml.ToString() );

      foreach( var element in xml.Descendants( XName.Get( "docPr", DocX.wp.NamespaceName ) ) )
        element.SetAttributeValue( XName.Get( "id" ), newDocPrId );

      foreach( var element in xml.Descendants( XName.Get( "blip", DocX.a.NamespaceName ) ) )
        element.SetAttributeValue( XName.Get( "embed", DocX.r.NamespaceName ), replaceWith.Id );

      var replacePicture = new Picture( document, xml, new Image( document, this.PackagePart.GetRelationship( replaceWith.Id ) ) );
      this.AppendPicture( replacePicture );
      toBeReplaced.Remove();

      return replacePicture;
    }

    /// <summary>
    /// Insert a Paragraph before this Paragraph, this Paragraph may have come from the same or another document.
    /// </summary>
    /// <param name="p">The Paragraph to insert.</param>
    /// <returns>The Paragraph now associated with this document.</returns>
    /// <example>
    /// Take a Paragraph from document a, and insert it into document b before this Paragraph.
    /// <code>
    /// // Place holder for a Paragraph.
    /// Paragraph p;
    ///
    /// // Load document a.
    /// using (DocX documentA = DocX.Load(@"a.docx"))
    /// {
    ///     // Get the first paragraph from this document.
    ///     p = documentA.Paragraphs[0];
    /// }
    ///
    /// // Load document b.
    /// using (DocX documentB = DocX.Load(@"b.docx"))
    /// {
    ///     // Get the first Paragraph in document b.
    ///     Paragraph p2 = documentB.Paragraphs[0];
    ///
    ///     // Insert the Paragraph from document a before this Paragraph.
    ///     Paragraph newParagraph = p2.InsertParagraphBeforeSelf(p);
    ///
    ///     // Save all changes made to document b.
    ///     documentB.Save();
    /// }// Release this document from memory.
    /// </code> 
    /// </example>
    public override Paragraph InsertParagraphBeforeSelf( Paragraph p )
    {
      var p2 = base.InsertParagraphBeforeSelf( p );
      p2.PackagePart = this.PackagePart;
      return p2;
    }

    /// <summary>
    /// Insert a new Paragraph before this Paragraph.
    /// </summary>
    /// <param name="text">The initial text for this new Paragraph.</param>
    /// <returns>A new Paragraph inserted before this Paragraph.</returns>
    /// <example>
    /// Insert a new paragraph before the first Paragraph in this document.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
    ///
    ///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.");
    ///
    ///     // Save all changes made to this new document.
    ///     document.Save();
    ///    }// Release this new document form memory.
    /// </code>
    /// </example>
    public override Paragraph InsertParagraphBeforeSelf( string text )
    {
      var p = base.InsertParagraphBeforeSelf( text );
      p.PackagePart = this.PackagePart;
      return p;
    }

    /// <summary>
    /// Insert a new Paragraph before this Paragraph.
    /// </summary>
    /// <param name="text">The initial text for this new Paragraph.</param>
    /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
    /// <returns>A new Paragraph inserted before this Paragraph.</returns>
    /// <example>
    /// Insert a new paragraph before the first Paragraph in this document.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
    ///
    ///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.", false);
    ///
    ///     // Save all changes made to this new document.
    ///     document.Save();
    ///    }// Release this new document form memory.
    /// </code>
    /// </example>
    public override Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges )
    {
      var p = base.InsertParagraphBeforeSelf( text, trackChanges );
      p.PackagePart = this.PackagePart;
      return p;
    }

    /// <summary>
    /// Insert a new Paragraph before this Paragraph.
    /// </summary>
    /// <param name="text">The initial text for this new Paragraph.</param>
    /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
    /// <param name="formatting">The formatting to apply to this insertion.</param>
    /// <returns>A new Paragraph inserted before this Paragraph.</returns>
    /// <example>
    /// Insert a new paragraph before the first Paragraph in this document.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
    ///
    ///     Formatting boldFormatting = new Formatting();
    ///     boldFormatting.Bold = true;
    ///
    ///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.", false, boldFormatting);
    ///
    ///     // Save all changes made to this new document.
    ///     document.Save();
    ///    }// Release this new document form memory.
    /// </code>
    /// </example>
    public override Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges, Formatting formatting )
    {
      var p = base.InsertParagraphBeforeSelf( text, trackChanges, formatting );
      p.PackagePart = this.PackagePart;
      return p;
    }

    /// <summary>
    /// Insert a page break before a Paragraph.
    /// </summary>
    /// <example>
    /// Insert 2 Paragraphs into a document with a page break between them.
    /// <code>
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///    // Insert a new Paragraph.
    ///    Paragraph p1 = document.InsertParagraph("Paragraph 1", false);
    ///       
    ///    // Insert a new Paragraph.
    ///    Paragraph p2 = document.InsertParagraph("Paragraph 2", false);
    ///    
    ///    // Insert a page break before Paragraph two.
    ///    p2.InsertPageBreakBeforeSelf();
    ///    
    ///    // Save this document.
    ///    document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public override void InsertPageBreakBeforeSelf()
    {
      base.InsertPageBreakBeforeSelf();
    }

    /// <summary>
    /// Insert a page break after a Paragraph.
    /// </summary>
    /// <example>
    /// Insert 2 Paragraphs into a document with a page break between them.
    /// <code>
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///    // Insert a new Paragraph.
    ///    Paragraph p1 = document.InsertParagraph("Paragraph 1", false);
    ///       
    ///    // Insert a page break after this Paragraph.
    ///    p1.InsertPageBreakAfterSelf();
    ///       
    ///    // Insert a new Paragraph.
    ///    Paragraph p2 = document.InsertParagraph("Paragraph 2", false);
    ///
    ///    // Save this document.
    ///    document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public override void InsertPageBreakAfterSelf()
    {
      base.InsertPageBreakAfterSelf();
    }

    [Obsolete( "Instead use: InsertHyperlink(Hyperlink h, int index)" )]
    public Paragraph InsertHyperlink( int index, Hyperlink h )
    {
      return InsertHyperlink( h, index );
    }

    /// <summary>
    /// This function inserts a hyperlink into a Paragraph at a specified character index.
    /// </summary>
    /// <param name="index">The index to insert at.</param>
    /// <param name="h">The hyperlink to insert.</param>
    /// <returns>The Paragraph with the Hyperlink inserted at the specified index.</returns>

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
      var Id = GetOrGenerateRel( h );

      XElement h_xml;
      if( index == 0 )
      {
        // Add this hyperlink as the first element.
        Xml.AddFirst( h.Xml );

        // Extract the picture back out of the DOM.
        h_xml = ( XElement )Xml.FirstNode;
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
          h_xml = ( XElement )Xml.LastNode;
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
          h_xml = ( XElement )run.Xml.NextNode;
        }

        h_xml.SetAttributeValue( DocX.r + "id", Id );
      }

      this._runs = Xml.Elements().Last().Elements( XName.Get( "r", DocX.w.NamespaceName ) ).ToList();
      return this;
    }

    /// <summary>
    /// Remove the Hyperlink at the provided index. The first hyperlink is at index 0.
    /// Using a negative index or an index greater than the index of the last hyperlink will cause an ArgumentOutOfRangeException() to be thrown.
    /// </summary>
    /// <param name="index">The index of the hyperlink to be removed.</param>
    /// <example>
    /// <code>
    /// // Crete a new document.
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///     // Add a Hyperlink into this document.
    ///     Hyperlink h = document.AddHyperlink("link", new Uri("http://www.google.com"));
    ///
    ///     // Insert a new Paragraph into the document.
    ///     Paragraph p1 = document.InsertParagraph("AC");
    ///     
    ///     // Insert the hyperlink into this Paragraph.
    ///     p1.InsertHyperlink(1, h);
    ///     Assert.IsTrue(p1.Text == "AlinkC"); // Make sure the hyperlink was inserted correctly;
    ///     
    ///     // Remove the hyperlink
    ///     p1.RemoveHyperlink(0);
    ///     Assert.IsTrue(p1.Text == "AC"); // Make sure the hyperlink was removed correctly;
    /// }
    /// </code>
    /// </example>
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

    /// <summary>
    /// Insert a Paragraph after this Paragraph, this Paragraph may have come from the same or another document.
    /// </summary>
    /// <param name="p">The Paragraph to insert.</param>
    /// <returns>The Paragraph now associated with this document.</returns>
    /// <example>
    /// Take a Paragraph from document a, and insert it into document b after this Paragraph.
    /// <code>
    /// // Place holder for a Paragraph.
    /// Paragraph p;
    ///
    /// // Load document a.
    /// using (DocX documentA = DocX.Load(@"a.docx"))
    /// {
    ///     // Get the first paragraph from this document.
    ///     p = documentA.Paragraphs[0];
    /// }
    ///
    /// // Load document b.
    /// using (DocX documentB = DocX.Load(@"b.docx"))
    /// {
    ///     // Get the first Paragraph in document b.
    ///     Paragraph p2 = documentB.Paragraphs[0];
    ///
    ///     // Insert the Paragraph from document a after this Paragraph.
    ///     Paragraph newParagraph = p2.InsertParagraphAfterSelf(p);
    ///
    ///     // Save all changes made to document b.
    ///     documentB.Save();
    /// }// Release this document from memory.
    /// </code> 
    /// </example>
    public override Paragraph InsertParagraphAfterSelf( Paragraph p )
    {
      var p2 = base.InsertParagraphAfterSelf( p );
      p2.PackagePart = this.PackagePart;
      return p2;
    }

    /// <summary>
    /// Insert a new Paragraph after this Paragraph.
    /// </summary>
    /// <param name="text">The initial text for this new Paragraph.</param>
    /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
    /// <param name="formatting">The formatting to apply to this insertion.</param>
    /// <returns>A new Paragraph inserted after this Paragraph.</returns>
    /// <example>
    /// Insert a new paragraph after the first Paragraph in this document.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
    ///
    ///     Formatting boldFormatting = new Formatting();
    ///     boldFormatting.Bold = true;
    ///
    ///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.", false, boldFormatting);
    ///
    ///     // Save all changes made to this new document.
    ///     document.Save();
    ///    }// Release this new document form memory.
    /// </code>
    /// </example>
    public override Paragraph InsertParagraphAfterSelf( string text, bool trackChanges, Formatting formatting )
    {
      var p = base.InsertParagraphAfterSelf( text, trackChanges, formatting );
      p.PackagePart = this.PackagePart;
      return p;
    }

    /// <summary>
    /// Insert a new Paragraph after this Paragraph.
    /// </summary>
    /// <param name="text">The initial text for this new Paragraph.</param>
    /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
    /// <returns>A new Paragraph inserted after this Paragraph.</returns>
    /// <example>
    /// Insert a new paragraph after the first Paragraph in this document.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
    ///
    ///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.", false);
    ///
    ///     // Save all changes made to this new document.
    ///     document.Save();
    ///    }// Release this new document form memory.
    /// </code>
    /// </example>
    public override Paragraph InsertParagraphAfterSelf( string text, bool trackChanges )
    {
      var p = base.InsertParagraphAfterSelf( text, trackChanges );
      p.PackagePart = this.PackagePart;
      return p;
    }

    /// <summary>
    /// Insert a new Paragraph after this Paragraph.
    /// </summary>
    /// <param name="text">The initial text for this new Paragraph.</param>
    /// <returns>A new Paragraph inserted after this Paragraph.</returns>
    /// <example>
    /// Insert a new paragraph after the first Paragraph in this document.
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
    ///
    ///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.");
    ///
    ///     // Save all changes made to this new document.
    ///     document.Save();
    ///    }// Release this new document form memory.
    /// </code>
    /// </example>
    public override Paragraph InsertParagraphAfterSelf( string text )
    {
      var p = base.InsertParagraphAfterSelf( text );
      p.PackagePart = this.PackagePart;
      return p;
    }

    /// <summary>
    /// Remove this Paragraph from the document.
    /// </summary>
    /// <param name="trackChanges">Should this remove be tracked as a change?</param>
    /// <example>
    /// Remove a Paragraph from a document and track it as a change.
    /// <code>
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Create and Insert a new Paragraph into this document.
    ///     Paragraph p = document.InsertParagraph("Hello", false);
    ///
    ///     // Remove the Paragraph and track this as a change.
    ///     p.Remove(true);
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public void Remove( bool trackChanges )
    {
      if( trackChanges )
      {
        DateTime now = DateTime.Now.ToUniversalTime();

        List<XElement> elements = Xml.Elements().ToList();
        List<XElement> temp = new List<XElement>();
        for( int i = 0; i < elements.Count(); i++ )
        {
          XElement e = elements[ i ];

          if( e.Name.LocalName != "del" )
          {
            temp.Add( e );
            e.Remove();
          }

          else
          {
            if( temp.Count() > 0 )
            {
              e.AddBeforeSelf( CreateEdit( EditType.del, now, temp.Elements() ) );
              temp.Clear();
            }
          }
        }

        if( temp.Count() > 0 )
          Xml.Add( CreateEdit( EditType.del, now, temp ) );
      }

      else
      {
        // If this is the only Paragraph in the Cell then we cannot remove it.
        if( Xml.Parent.Name.LocalName == "tc" && Xml.Parent.Elements( XName.Get( "p", DocX.w.NamespaceName ) ).Count() == 1 )
          Xml.Value = string.Empty;

        else
        {
          // Remove this paragraph from the document
          Xml.Remove();
          Xml = null;
        }
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
    // using (DocX document = DocX.Create(@"Test.docx"))
    // {
    //     // Add a new Paragraph to this document.
    //     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
    //
    //     // Add an Image to this document.
    //     Xceed.Words.NET.Image img = document.AddImage(@"Image.jpg");
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
    //    DocX.RenumberIDs(document);
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
    // using (DocX document = DocX.Create(@"Test.docx"))
    // {
    //     // Add a new Paragraph to this document.
    //     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
    //
    //     // Add an Image to this document.
    //     Xceed.Words.NET.Image img = document.AddImage(@"Image.jpg");
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

    /// <summary>
    /// Inserts a specified instance of System.String into a Xceed.Words.NET.DocX.Paragraph at a specified index position.
    /// </summary>
    /// <example>
    /// <code> 
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Create a text formatting.
    ///     Formatting f = new Formatting();
    ///     f.FontColor = Color.Red;
    ///     f.Size = 30;
    ///
    ///     // Iterate through the Paragraphs in this document.
    ///     foreach (Paragraph p in document.Paragraphs)
    ///     {
    ///         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
    ///         p.InsertText("Start: ", true, f);
    ///     }
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <example>
    /// Inserting tabs using the \t switch.
    /// <code>  
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///      // Create a text formatting.
    ///      Formatting f = new Formatting();
    ///      f.FontColor = Color.Red;
    ///      f.Size = 30;
    ///        
    ///      // Iterate through the paragraphs in this document.
    ///      foreach (Paragraph p in document.Paragraphs)
    ///      {
    ///          // Insert the string "\tEnd" at the end of every paragraph and flag it as a change.
    ///          p.InsertText("\tEnd", true, f);
    ///      }
    ///       
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
    /// <seealso cref="Paragraph.RemoveText(int, int, bool, bool)"/>
    /// <param name="value">The System.String to insert.</param>
    /// <param name="trackChanges">Flag this insert as a change.</param>
    /// <param name="formatting">The text formatting.</param>
    public void InsertText( string value, bool trackChanges = false, Formatting formatting = null )
    {
      this.InsertText( this.Text.Length, value, trackChanges, formatting );
    }

    /// <summary>
    /// Inserts a specified instance of System.String into a Xceed.Words.NET.DocX.Paragraph at a specified index position.
    /// </summary>
    /// <example>
    /// <code> 
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Create a text formatting.
    ///     Formatting f = new Formatting();
    ///     f.FontColor = Color.Red;
    ///     f.Size = 30;
    ///
    ///     // Iterate through the Paragraphs in this document.
    ///     foreach (Paragraph p in document.Paragraphs)
    ///     {
    ///         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
    ///         p.InsertText(0, "Start: ", true, f);
    ///     }
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <example>
    /// Inserting tabs using the \t switch.
    /// <code>  
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Create a text formatting.
    ///     Formatting f = new Formatting();
    ///     f.FontColor = Color.Red;
    ///     f.Size = 30;
    ///
    ///     // Iterate through the paragraphs in this document.
    ///     foreach (Paragraph p in document.Paragraphs)
    ///     {
    ///         // Insert the string "\tStart:\t" at the begining of every paragraph and flag it as a change.
    ///         p.InsertText(0, "\tStart:\t", true, f);
    ///     }
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
    /// <seealso cref="Paragraph.RemoveText(int, int, bool, bool)"/>
    /// <param name="index">The index position of the insertion.</param>
    /// <param name="value">The System.String to insert.</param>
    /// <param name="trackChanges">Flag this insert as a change.</param>
    /// <param name="formatting">The text formatting.</param>
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
          insert = CreateEdit( EditType.ins, insert_datetime, insert );
        }
        this.Xml.Add( insert );
      }
      else
      {
        object newRuns = null;
        var rPr = run.Xml.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );

        if( formatting != null )
        {
          Formatting oldFormatting = null;
          Formatting newFormatting = null;

          if( rPr != null )
          {
            oldFormatting = Formatting.Parse( rPr );
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
              var parent_ins_date = DateTime.Parse( parentElement.Attribute( XName.Get( "date", DocX.w.NamespaceName ) ).Value );

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
                insert = CreateEdit( EditType.ins, insert_datetime, newRuns );
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
                insert = CreateEdit( EditType.ins, insert_datetime, newRuns );
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

      HelperFunctions.RenumberIDs( Document );
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="culture">The CultureInfo for text</param>
    /// <returns>This Paragraph in curent culture</returns>
    /// <example>
    /// Add a new Paragraph with russian text to this document and then set language of text to local culture.
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph with russian text and set specific culture to it.
    ///     Paragraph p = document.InsertParagraph("Привет мир").Culture(CultureInfo.CreateSpecificCulture("ru-RU"));
    ///       
    ///     // Save this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public Paragraph Culture( CultureInfo culture )
    {
      this.ApplyTextFormattingProperty( XName.Get( "lang", DocX.w.NamespaceName ),
                                        string.Empty,
                                        new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), culture.Name ) );
      return this;
    }

    /// <summary>
    /// Append text to this Paragraph.
    /// </summary>
    /// <param name="text">The text to append.</param>
    /// <returns>This Paragraph with the new text appened.</returns>
    /// <example>
    /// Add a new Paragraph to this document and then append some text to it.
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph and Append some text to it.
    ///     Paragraph p = document.InsertParagraph().Append("Hello World!!!");
    ///       
    ///     // Save this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public Paragraph Append( string text )
    {
      List<XElement> newRuns = HelperFunctions.FormatInput( text, null );
      Xml.Add( newRuns );

      _runs = Xml.Elements( XName.Get( "r", DocX.w.NamespaceName ) ).Reverse().Take( newRuns.Count() ).ToList();

      return this;
    }

    /// <summary>
    /// Append text to this Paragraph and apply the provided format
    /// </summary>
    /// <param name="text">The text to append.</param>
    /// <param name="format">The format to use.</param>
    /// <returns>This Paragraph with the new text appended.</returns>
    /// <example>
    /// Add a new Paragraph to this document, append some text to it and apply the provided format.
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Prepare format to use
    ///     Formatting format = new Formatting();
    ///     format.Bold = true;
    ///     format.Size = 18;
    ///     format.FontColor = Color.Blue;
    /// 
    ///     // Insert a new Paragraph and append some text to it with the custom format
    ///     Paragraph p = document.InsertParagraph().Append("Hello World!!!", format);
    ///       
    ///     // Save this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public Paragraph Append( string text, Formatting format )
    {
      // Text
      Append( text );

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

    /// <summary>
    /// Append a hyperlink to a Paragraph.
    /// </summary>
    /// <param name="h">The hyperlink to append.</param>
    /// <returns>The Paragraph with the hyperlink appended.</returns>
    /// <example>
    /// Creates a Paragraph with some text and a hyperlink.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///    // Add a hyperlink to this document.
    ///    Hyperlink h = document.AddHyperlink("Google", new Uri("http://www.google.com"));
    ///    
    ///    // Add a new Paragraph to this document.
    ///    Paragraph p = document.InsertParagraph();
    ///    p.Append("My favourite search engine is ");
    ///    p.AppendHyperlink(h);
    ///    p.Append(", I think it's great.");
    ///
    ///    // Save all changes made to this document.
    ///    document.Save();
    /// }
    /// </code>
    /// </example>
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
      var Id = GetOrGenerateRel( h );

      Xml.Add( h.Xml );
      Xml.Elements().Last().SetAttributeValue( DocX.r + "id", Id );

      _runs = Xml.Elements().Last().Elements( XName.Get( "r", DocX.w.NamespaceName ) ).ToList();

      return this;
    }

    /// <summary>
    /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
    /// </summary>
    /// <param name="p">The Picture to append.</param>
    /// <returns>The Paragraph with the Picture now appended.</returns>
    /// <example>
    /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
    /// <code>
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///    // Add an image to the document. 
    ///    Image     i = document.AddImage(@"Image.jpg");
    ///    
    ///    // Create a picture i.e. (A custom view of an image)
    ///    Picture   p = i.CreatePicture();
    ///    p.FlipHorizontal = true;
    ///    p.Rotation = 10;
    ///
    ///    // Create a new Paragraph.
    ///    Paragraph par = document.InsertParagraph();
    ///    
    ///    // Append content to the Paragraph.
    ///    par.Append("Here is a cool picture")
    ///       .AppendPicture(p)
    ///       .Append(" don't you think so?");
    ///
    ///    // Save all changes made to this document.
    ///    document.Save();
    /// }
    /// </code>
    /// </example>
    public Paragraph AppendPicture( Picture p )
    {
      // Convert the path of this mainPart to its equilivant rels file path.
      var path = this.PackagePart.Uri.OriginalString.Replace( "/word/", "" );
      var rels_path = new Uri( "/word/_rels/" + path + ".rels", UriKind.Relative );

      // Check to see if the rels file exists and create it if not.
      if( !Document._package.PartExists( rels_path ) )
      {
        HelperFunctions.CreateRelsPackagePart( Document, rels_path );
      }

      // Check to see if a rel for this Picture exists, create it if not.
      var Id = GetOrGenerateRel( p );

      // Add the Picture Xml to the end of the Paragragraph Xml.
      Xml.Add( p.Xml );

      // Extract the attribute id from the Pictures Xml.
      var a_id =
      (
          from e in Xml.Elements().Last().Descendants()
          where e.Name.LocalName.Equals( "blip" )
          select e.Attribute( XName.Get( "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) )
      ).Single();

      // Set its value to the Pictures relationships id.
      a_id.SetValue( Id );

      // For formatting such as .Bold()
      _runs = Xml.Elements( XName.Get( "r", DocX.w.NamespaceName ) ).Reverse().Take( p.Xml.Elements( XName.Get( "r", DocX.w.NamespaceName ) ).Count() ).ToList();

      return this;
    }

    /// <summary>
    /// Add an equation to a document.
    /// </summary>
    /// <param name="equation">The Equation to append.</param>
    /// <returns>The Paragraph with the Equation now appended.</returns>
    /// <example>
    /// Add an equation to a document.
    /// <code>
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///    // Add an equation to the document. 
    ///    document.AddEquation("x=y+z");
    ///    
    ///    // Save all changes made to this document.
    ///    document.Save();
    /// }
    /// </code>
    /// </example>
    public Paragraph AppendEquation( String equation )
    {
      // Create equation element
      XElement oMathPara =
          new XElement
          (
              XName.Get( "oMathPara", DocX.m.NamespaceName ),
              new XElement
              (
                  XName.Get( "oMath", DocX.m.NamespaceName ),
                  new XElement
                  (
                      XName.Get( "r", DocX.w.NamespaceName ),
                      new Formatting() { FontFamily = new Font( "Cambria Math" ) }.Xml,                           // create formatting
                      new XElement( XName.Get( "t", DocX.m.NamespaceName ), equation )                            // create equation string
                  )
              )
          );

      // Add equation element into paragraph xml and update runs collection
      Xml.Add( oMathPara );
      _runs = Xml.Elements( XName.Get( "oMathPara", DocX.m.NamespaceName ) ).ToList();

      // Return paragraph with equation
      return this;
    }

    /// <summary>
    /// Insert a Picture into a Paragraph at the given text index.
    /// If not index is provided defaults to 0.
    /// </summary>
    /// <param name="p">The Picture to insert.</param>
    /// <param name="index">The text index to insert at.</param>
    /// <returns>The modified Paragraph.</returns>
    /// <example>
    /// <code>
    ///Load test document.
    ///using (DocX document = DocX.Create("Test.docx"))
    ///{
    ///    // Add Headers and Footers into this document.
    ///    document.AddHeaders();
    ///    document.AddFooters();
    ///    document.DifferentFirstPage = true;
    ///    document.DifferentOddAndEvenPages = true;
    ///
    ///    // Add an Image to this document.
    ///    Xceed.Words.NET.Image img = document.AddImage(directory_documents + "purple.png");
    ///
    ///    // Create a Picture from this Image.
    ///    Picture pic = img.CreatePicture();
    ///
    ///    // Main document.
    ///    Paragraph p0 = document.InsertParagraph("Hello");
    ///    p0.InsertPicture(pic, 3);
    ///
    ///    // Header first.
    ///    Paragraph p1 = document.Headers.first.InsertParagraph("----");
    ///    p1.InsertPicture(pic, 2);
    ///
    ///    // Header odd.
    ///    Paragraph p2 = document.Headers.odd.InsertParagraph("----");
    ///    p2.InsertPicture(pic, 2);
    ///
    ///    // Header even.
    ///    Paragraph p3 = document.Headers.even.InsertParagraph("----");
    ///    p3.InsertPicture(pic, 2);
    ///
    ///    // Footer first.
    ///    Paragraph p4 = document.Footers.first.InsertParagraph("----");
    ///    p4.InsertPicture(pic, 2);
    ///
    ///    // Footer odd.
    ///    Paragraph p5 = document.Footers.odd.InsertParagraph("----");
    ///    p5.InsertPicture(pic, 2);
    ///
    ///    // Footer even.
    ///    Paragraph p6 = document.Footers.even.InsertParagraph("----");
    ///    p6.InsertPicture(pic, 2);
    ///
    ///    // Save this document.
    ///    document.Save();
    ///}
    /// </code>
    /// </example>
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

      // Check to see if a rel for this Picture exists, create it if not.
      var Id = GetOrGenerateRel( p );

      XElement p_xml;
      if( index == 0 )
      {
        // Add this hyperlink as the last element.
        Xml.AddFirst( p.Xml );

        // Extract the picture back out of the DOM.
        p_xml = ( XElement )Xml.FirstNode;
      }
      else
      {
        // Get the first run effected by this Insert
        var run = GetFirstRunEffectedByEdit( index );
        if( run == null )
        {
          // Add this picture as the last element.
          Xml.Add( p.Xml );

          // Extract the picture back out of the DOM.
          p_xml = ( XElement )Xml.LastNode;
        }
        else
        {
          // Split this run at the point you want to insert
          var splitRun = Run.SplitRun( run, index );

          // Replace the origional run.
          run.Xml.ReplaceWith( splitRun[ 0 ], p.Xml, splitRun[ 1 ] );

          // Get the first run effected by this Insert
          run = GetFirstRunEffectedByEdit( index );

          // The picture has to be the next element, extract it back out of the DOM.
          p_xml = ( XElement )run.Xml.NextNode;
        }
      }

      // Extract the attribute id from the Pictures Xml.
      XAttribute a_id =
      (
          from e in p_xml.Descendants()
          where e.Name.LocalName.Equals( "blip" )
          select e.Attribute( XName.Get( "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) )
      ).Single();

      // Set its value to the Pictures relationships id.
      a_id.SetValue( Id );

      return this;
    }

    /// <summary>
    /// Add a new TabStopPosition in the current paragraph.
    /// </summary>
    /// <param name="alignment">Specifies the alignment of the Tab stop.</param>
    /// <param name="position">Specifies the horizontal position of the tab stop.</param>
    /// <param name="leader">Specifies the character used to fill in the space created by a tab.</param>
    /// <returns>The modified Paragraph.</returns>
    public Paragraph InsertTabStopPosition( Alignment alignment, float position, TabStopPositionLeader leader = TabStopPositionLeader.none )
    {
      var pPr = GetOrCreate_pPr();
      var tabs = pPr.Element( XName.Get( "tabs", DocX.w.NamespaceName ) );
      if( tabs == null )
      {
        tabs = new XElement( XName.Get( "tabs", DocX.w.NamespaceName ) );
        pPr.Add( tabs );
      }

      var newTab = new XElement( XName.Get( "tab", DocX.w.NamespaceName ) );

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
      newTab.SetAttributeValue( XName.Get( "val", DocX.w.NamespaceName ), alignmentString );

      // Position
      var posValue = position * 20.0f;
      newTab.SetAttributeValue( XName.Get( "pos", DocX.w.NamespaceName ), posValue.ToString() );

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
      newTab.SetAttributeValue( XName.Get( "leader", DocX.w.NamespaceName ), leaderString );

      tabs.Add( newTab );
      return this;
    }

    /// <summary>
    /// Append text on a new line to this Paragraph.
    /// </summary>
    /// <param name="text">The text to append.</param>
    /// <returns>This Paragraph with the new text appened.</returns>
    /// <example>
    /// Add a new Paragraph to this document and then append a new line with some text to it.
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph and Append a new line with some text to it.
    ///     Paragraph p = document.InsertParagraph().AppendLine("Hello World!!!");
    ///       
    ///     // Save this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public Paragraph AppendLine( string text )
    {
      return Append( "\n" + text );
    }

    /// <summary>
    /// Append a new line to this Paragraph.
    /// </summary>
    /// <returns>This Paragraph with a new line appeneded.</returns>
    /// <example>
    /// Add a new Paragraph to this document and then append a new line to it.
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph and Append a new line with some text to it.
    ///     Paragraph p = document.InsertParagraph().AppendLine();
    ///       
    ///     // Save this document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    public Paragraph AppendLine()
    {
      return Append( "\n" );
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <returns>This Paragraph with the last appended text bold.</returns>
    /// <example>
    /// Append text to this Paragraph and then make it bold.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("Bold").Bold()
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph Bold()
    {
      ApplyTextFormattingProperty( XName.Get( "b", DocX.w.NamespaceName ), string.Empty, null );
      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <returns>This Paragraph with the last appended text italic.</returns>
    /// <example>
    /// Append text to this Paragraph and then make it italic.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("Italic").Italic()
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph Italic()
    {
      ApplyTextFormattingProperty( XName.Get( "i", DocX.w.NamespaceName ), string.Empty, null );
      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="c">A color to use on the appended text.</param>
    /// <returns>This Paragraph with the last appended text colored.</returns>
    /// <example>
    /// Append text to this Paragraph and then color it.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("Blue").Color(Color.Blue)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph Color( Color c )
    {
      ApplyTextFormattingProperty( XName.Get( "color", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), c.ToHex() ) );
      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="underlineStyle">The underline style to use for the appended text.</param>
    /// <returns>This Paragraph with the last appended text underlined.</returns>
    /// <example>
    /// Append text to this Paragraph and then underline it.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("Underlined").UnderlineStyle(UnderlineStyle.doubleLine)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph UnderlineStyle( UnderlineStyle underlineStyle )
    {
      string value;
      switch( underlineStyle )
      {
        case Xceed.Words.NET.UnderlineStyle.none:
          value = string.Empty;
          break;
        case Xceed.Words.NET.UnderlineStyle.singleLine:
          value = "single";
          break;
        case Xceed.Words.NET.UnderlineStyle.doubleLine:
          value = "double";
          break;
        default:
          value = underlineStyle.ToString();
          break;
      }

      ApplyTextFormattingProperty( XName.Get( "u", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), value ) );
      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="fontSize">The font size to use for the appended text.</param>
    /// <returns>This Paragraph with the last appended text resized.</returns>
    /// <example>
    /// Append text to this Paragraph and then resize it.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("Big").FontSize(20)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph FontSize( double fontSize )
    {
      double tempSize = (int)fontSize*2;
      if( tempSize - ( int )tempSize == 0 )
      {
        if( !( fontSize > 0 && fontSize < 1639 ) )
          throw new ArgumentException( "Size", "Value must be in the range 0 - 1638" );
      }

      else
        throw new ArgumentException( "Size", "Value must be either a whole or half number, examples: 32, 32.5" );

      ApplyTextFormattingProperty( XName.Get( "sz", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), fontSize * 2 ) );
      ApplyTextFormattingProperty( XName.Get( "szCs", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), fontSize * 2 ) );

      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="fontName">The font to use for the appended text.</param>
    /// <returns>This Paragraph with the last appended text's font changed.</returns>
    public Paragraph Font (string fontName)
    {
      return Font(new Font(fontName));
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="fontFamily">The font to use for the appended text.</param>
    /// <returns>This Paragraph with the last appended text's font changed.</returns>
    /// <example>
    /// Append text to this Paragraph and then change its font.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("Times new roman").Font(new FontFamily("Times new roman"))
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph Font( Font fontFamily )
    {
      ApplyTextFormattingProperty
      (
          XName.Get( "rFonts", DocX.w.NamespaceName ),
          string.Empty,
          new[]
          {
            new XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily.Name),
            new XAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), fontFamily.Name),
            new XAttribute(XName.Get("cs", DocX.w.NamespaceName), fontFamily.Name),
            new XAttribute(XName.Get("eastAsia", DocX.w.NamespaceName), fontFamily.Name),
          }
      );

      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="capsStyle">The caps style to apply to the last appended text.</param>
    /// <returns>This Paragraph with the last appended text's caps style changed.</returns>
    /// <example>
    /// Append text to this Paragraph and then set it to full caps.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("Capitalized").CapsStyle(CapsStyle.caps)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph CapsStyle( CapsStyle capsStyle )
    {
      switch( capsStyle )
      {
        case Xceed.Words.NET.CapsStyle.none:
          break;

        default:
          {
            ApplyTextFormattingProperty( XName.Get( capsStyle.ToString(), DocX.w.NamespaceName ), string.Empty, null );
            break;
          }
      }

      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="script">The script style to apply to the last appended text.</param>
    /// <returns>This Paragraph with the last appended text's script style changed.</returns>
    /// <example>
    /// Append text to this Paragraph and then set it to superscript.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("superscript").Script(Script.superscript)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph Script( Script script )
    {
      switch( script )
      {
        case Xceed.Words.NET.Script.none:
          break;

        default:
          {
            ApplyTextFormattingProperty( XName.Get( "vertAlign", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), script.ToString() ) );
            break;
          }
      }

      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    ///<param name="highlight">The highlight to apply to the last appended text.</param>
    /// <returns>This Paragraph with the last appended text highlighted.</returns>
    /// <example>
    /// Append text to this Paragraph and then highlight it.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("highlighted").Highlight(Highlight.green)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph Highlight( Highlight highlight )
    {
      switch( highlight )
      {
        case Xceed.Words.NET.Highlight.none:
          break;

        default:
          {
            ApplyTextFormattingProperty( XName.Get( "highlight", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), highlight.ToString() ) );
            break;
          }
      }

      return this;
    }

    public Paragraph Shading( Color shading, ShadingType shadingType = ShadingType.Text )
    {
      // Add to run
      if( shadingType == ShadingType.Text )
      {
        this.ApplyTextFormattingProperty( XName.Get( "shd", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "fill", DocX.w.NamespaceName ), shading.ToHex() ) );
      }
      // Add to paragraph
      else
      {
        var pPr = GetOrCreate_pPr();
        var shd = pPr.Element( XName.Get( "shd", DocX.w.NamespaceName ) );
        if( shd == null )
        {
          shd = new XElement( XName.Get( "shd", DocX.w.NamespaceName ) );
          pPr.Add( shd );
        }

        var fillAttribute = shd.Attribute( XName.Get( "fill", DocX.w.NamespaceName ) );
        if( fillAttribute == null )
        {
          shd.SetAttributeValue( XName.Get( "fill", DocX.w.NamespaceName ), shading.ToHex() );
        }
        else
        {
          fillAttribute.SetValue( shading.ToHex() );
        }
      }

      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="misc">The miscellaneous property to set.</param>
    /// <returns>This Paragraph with the last appended text changed by a miscellaneous property.</returns>
    /// <example>
    /// Append text to this Paragraph and then apply a miscellaneous property.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("outlined").Misc(Misc.outline)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph Misc( Misc misc )
    {
      switch( misc )
      {
        case Xceed.Words.NET.Misc.none:
          break;

        case Xceed.Words.NET.Misc.outlineShadow:
          {
            ApplyTextFormattingProperty( XName.Get( "outline", DocX.w.NamespaceName ), string.Empty, null );
            ApplyTextFormattingProperty( XName.Get( "shadow", DocX.w.NamespaceName ), string.Empty, null );

            break;
          }

        case Xceed.Words.NET.Misc.engrave:
          {
            ApplyTextFormattingProperty( XName.Get( "imprint", DocX.w.NamespaceName ), string.Empty, null );

            break;
          }

        default:
          {
            ApplyTextFormattingProperty( XName.Get( misc.ToString(), DocX.w.NamespaceName ), string.Empty, null );

            break;
          }
      }

      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="strikeThrough">The strike through style to used on the last appended text.</param>
    /// <returns>This Paragraph with the last appended text striked.</returns>
    /// <example>
    /// Append text to this Paragraph and then strike it.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("striked").StrikeThrough(StrikeThrough.doubleStrike)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph StrikeThrough( StrikeThrough strikeThrough )
    {
      string value;
      switch( strikeThrough )
      {
        case Xceed.Words.NET.StrikeThrough.strike:
          value = "strike";
          break;
        case Xceed.Words.NET.StrikeThrough.doubleStrike:
          value = "dstrike";
          break;
        default:
          return this;
      }

      ApplyTextFormattingProperty( XName.Get( value, DocX.w.NamespaceName ), string.Empty, null );

      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <param name="underlineColor">The underline color to use, if no underline is set, a single line will be used.</param>
    /// <returns>This Paragraph with the last appended text underlined in a color.</returns>
    /// <example>
    /// Append text to this Paragraph and then underline it using a color.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("color underlined").UnderlineStyle(UnderlineStyle.dotted).UnderlineColor(Color.Orange)
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph UnderlineColor( Color underlineColor )
    {
      foreach( XElement run in _runs )
      {
        XElement rPr = run.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );
        if( rPr == null )
        {
          run.AddFirst( new XElement( XName.Get( "rPr", DocX.w.NamespaceName ) ) );
          rPr = run.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );
        }

        XElement u = rPr.Element( XName.Get( "u", DocX.w.NamespaceName ) );
        if( u == null )
        {
          rPr.SetElementValue( XName.Get( "u", DocX.w.NamespaceName ), string.Empty );
          u = rPr.Element( XName.Get( "u", DocX.w.NamespaceName ) );
          u.SetAttributeValue( XName.Get( "val", DocX.w.NamespaceName ), "single" );
        }

        u.SetAttributeValue( XName.Get( "color", DocX.w.NamespaceName ), underlineColor.ToHex() );
      }

      return this;
    }

    /// <summary>
    /// For use with Append() and AppendLine()
    /// </summary>
    /// <returns>This Paragraph with the last appended text hidden.</returns>
    /// <example>
    /// Append text to this Paragraph and then hide it.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Insert a new Paragraph.
    ///     Paragraph p = document.InsertParagraph();
    ///
    ///     p.Append("I am ")
    ///     .Append("hidden").Hide()
    ///     .Append(" I am not");
    ///        
    ///     // Save this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public Paragraph Hide()
    {
      ApplyTextFormattingProperty( XName.Get( "vanish", DocX.w.NamespaceName ), string.Empty, null );

      return this;
    }

    public Paragraph Spacing( double spacing )
    {
      spacing *= 20;

      if( spacing - ( int )spacing == 0 )
      {
        if( !( spacing > -1585 && spacing < 1585 ) )
          throw new ArgumentException( "Spacing", "Value must be in the range: -1584 - 1584" );
      }

      else
        throw new ArgumentException( "Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9" );

      ApplyTextFormattingProperty( XName.Get( "spacing", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), spacing ) );

      return this;
    }

    public Paragraph SpacingBefore( double spacingBefore )
    {
      spacingBefore *= 20;

      var pPr = GetOrCreate_pPr();
      var spacing = pPr.Element( XName.Get( "spacing", DocX.w.NamespaceName ) );

      if( spacingBefore > 0 )
      {
        if( spacing == null )
        {
          spacing = new XElement( XName.Get( "spacing", DocX.w.NamespaceName ) );
          pPr.Add( spacing );
        }

        var beforeAttribute = spacing.Attribute( XName.Get( "before", DocX.w.NamespaceName ) );

        if( beforeAttribute == null )
          spacing.SetAttributeValue( XName.Get( "before", DocX.w.NamespaceName ), spacingBefore );
        else
          beforeAttribute.SetValue( spacingBefore );
      }

      if( Math.Abs( spacingBefore ) < 0.1f && spacing != null )
      {
        var beforeAttribute = spacing.Attribute( XName.Get( "before", DocX.w.NamespaceName ) );
        beforeAttribute.Remove();

        if( !spacing.HasAttributes )
          spacing.Remove();
      }
      return this;
    }

    public Paragraph SpacingAfter( double spacingAfter )
    {
      spacingAfter *= 20;

      var pPr = GetOrCreate_pPr();
      var spacing = pPr.Element( XName.Get( "spacing", DocX.w.NamespaceName ) );

      if( spacingAfter > 0 )
      {
        if( spacing == null )
        {
          spacing = new XElement( XName.Get( "spacing", DocX.w.NamespaceName ) );
          pPr.Add( spacing );
        }

        var afterAttribute = spacing.Attribute( XName.Get( "after", DocX.w.NamespaceName ) );

        if( afterAttribute == null )
          spacing.SetAttributeValue( XName.Get( "after", DocX.w.NamespaceName ), spacingAfter );
        else
          afterAttribute.SetValue( spacingAfter );
      }

      if( Math.Abs( spacingAfter ) < 0.1f && spacing != null )
      {
        var afterAttribute = spacing.Attribute( XName.Get( "after", DocX.w.NamespaceName ) );
        afterAttribute.Remove();

        if( !spacing.HasAttributes )
          spacing.Remove();
      }
      return this;
    }

    public Paragraph Kerning( int kerning )
    {
      if( !new int?[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 }.Contains( kerning ) )
        throw new ArgumentOutOfRangeException( "Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72" );

      ApplyTextFormattingProperty( XName.Get( "kern", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), kerning * 2 ) );
      return this;
    }

    public Paragraph Position( double position )
    {
      if( !( position > -1585 && position < 1585 ) )
        throw new ArgumentOutOfRangeException( "Position", "Value must be in the range -1585 - 1585" );

      ApplyTextFormattingProperty( XName.Get( "position", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), position * 2 ) );

      return this;
    }

    public Paragraph PercentageScale( int percentageScale )
    {
      if( !( new int?[] { 200, 150, 100, 90, 80, 66, 50, 33 } ).Contains( percentageScale ) )
        throw new ArgumentOutOfRangeException( "PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33" );

      ApplyTextFormattingProperty( XName.Get( "w", DocX.w.NamespaceName ), string.Empty, new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), percentageScale ) );

      return this;
    }

    /// <summary>
    /// Append a field of type document property, this field will display the custom property cp, at the end of this paragraph.
    /// </summary>
    /// <param name="cp">The custom property to display.</param>
    /// <param name="f">The formatting to use for this text.</param>
    /// <param name="trackChanges"></param>
    /// <example>
    /// Create, add and display a custom property in a document.
    /// <code>
    /// // Load a document.
    ///using (DocX document = DocX.Create("CustomProperty_Add.docx"))
    ///{
    ///    // Add a few Custom Properties to this document.
    ///    document.AddCustomProperty(new CustomProperty("fname", "cathal"));
    ///    document.AddCustomProperty(new CustomProperty("age", 24));
    ///    document.AddCustomProperty(new CustomProperty("male", true));
    ///    document.AddCustomProperty(new CustomProperty("newyear2012", new DateTime(2012, 1, 1)));
    ///    document.AddCustomProperty(new CustomProperty("fav_num", 3.141592));
    ///
    ///    // Insert a new Paragraph and append a load of DocProperties.
    ///    Paragraph p = document.InsertParagraph("fname: ")
    ///        .AppendDocProperty(document.CustomProperties["fname"])
    ///        .Append(", age: ")
    ///        .AppendDocProperty(document.CustomProperties["age"])
    ///        .Append(", male: ")
    ///        .AppendDocProperty(document.CustomProperties["male"])
    ///        .Append(", newyear2012: ")
    ///        .AppendDocProperty(document.CustomProperties["newyear2012"])
    ///        .Append(", fav_num: ")
    ///        .AppendDocProperty(document.CustomProperties["fav_num"]);
    ///    
    ///    // Save the changes to the document.
    ///    document.Save();
    ///}
    /// </code>
    /// </example>
    public Paragraph AppendDocProperty( CustomProperty cp, bool trackChanges = false, Formatting f = null )
    {
      this.InsertDocProperty( cp, trackChanges, f );
      return this;
    }

    /// <summary>
    /// Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
    /// </summary>
    /// <param name="cp">The custom property to display.</param>
    /// <param name="trackChanges">if the changes are tracked.</param>
    /// <param name="f">The formatting to use for this text.</param>
    /// <example>
    /// Create, add and display a custom property in a document.
    /// <code>
    /// // Load a document
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Create a custom property.
    ///     CustomProperty name = new CustomProperty("name", "Cathal Coffey");
    ///        
    ///     // Add this custom property to this document.
    ///     document.AddCustomProperty(name);
    ///
    ///     // Create a text formatting.
    ///     Formatting f = new Formatting();
    ///     f.Bold = true;
    ///     f.Size = 14;
    ///     f.StrikeThrough = StrickThrough.strike;
    ///
    ///     // Insert a new paragraph.
    ///     Paragraph p = document.InsertParagraph("Author: ", false, f);
    ///
    ///     // Insert a field of type document property to display the custom property name and track this change.
    ///     p.InsertDocProperty(name, true, f);
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public DocProperty InsertDocProperty( CustomProperty cp, bool trackChanges = false, Formatting f = null )
    {
      XElement f_xml = null;
      if( f != null )
      {
        f_xml = f.Xml;
      }

      var e = new XElement( XName.Get( "fldSimple", DocX.w.NamespaceName ),
                            new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), string.Format( @"DOCPROPERTY {0} \* MERGEFORMAT", cp.Name ) ),
                            new XElement( XName.Get( "r", DocX.w.NamespaceName ), new XElement( XName.Get( "t", DocX.w.NamespaceName ), f_xml, cp.Value ) )
      );

      var xml = e;
      if( trackChanges )
      {
        var now = DateTime.Now;
        var insert_datetime = new DateTime( now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc );
        e = CreateEdit( EditType.ins, insert_datetime, e );
      }

      this.Xml.Add( e );

      return new DocProperty( this.Document, xml );
    }

    /// <summary>
    /// Removes characters from a Xceed.Words.NET.DocX.Paragraph.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Iterate through the paragraphs
    ///     foreach (Paragraph p in document.Paragraphs)
    ///     {
    ///         // Remove the first two characters from every paragraph
    ///         p.RemoveText(0, 2, false);
    ///     }
    ///        
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
    /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
    /// <param name="index">The position to begin deleting characters.</param>
    /// <param name="count">The number of characters to delete</param>
    /// <param name="trackChanges">Track changes</param>
    /// <param name="removeEmptyParagraph">Remove empty paragraph</param>
    public void RemoveText( int index, int count, bool trackChanges = false, bool removeEmptyParagraph = true )
    {
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
              var middle = Paragraph.CreateEdit( EditType.del, remove_datetime, temp.Elements() );
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
              var splitRunBefore = Run.SplitRun( run, index, EditType.del );
              var min = Math.Min( index + ( count - processed ), run.EndIndex );
              var splitRunAfter = Run.SplitRun( run, min, EditType.del );

              var middle = Paragraph.CreateEdit( EditType.del, remove_datetime, new List<XElement>() { Run.SplitRun( new Run( Document, splitRunBefore[ 1 ], run.StartIndex + GetElementTextLength( splitRunBefore[ 0 ] ) ), min, EditType.del )[ 0 ] } );
              processed += Paragraph.GetElementTextLength( middle as XElement );

              if( !trackChanges )
              {
                middle = null;
              }

              run.Xml.ReplaceWith( splitRunBefore[ 0 ], middle, splitRunAfter[ 1 ] );
              break;
            }
        }

        // In some cases, removing an empty paragraph is allowed
        var canRemove = removeEmptyParagraph && GetElementTextLength( parentElement ) == 0;
        if( parentElement.Parent != null )
        {
          // Need to make sure there is another paragraph in the parent cell
          canRemove &= parentElement.Parent.Name.LocalName == "tc" && parentElement.Parent.Elements( XName.Get( "p", DocX.w.NamespaceName ) ).Count() > 1;

          // Need to make sure there is no drawing element within the parent element.
          // Picture elements contain no text length but they are still content.
          canRemove &= parentElement.Descendants( XName.Get( "drawing", DocX.w.NamespaceName ) ).Count() == 0;

          if( canRemove )
            parentElement.Remove();
        }
      }
      while( processed < count );

      HelperFunctions.RenumberIDs( Document );
    }


    /// <summary>
    /// Removes characters from a Xceed.Words.NET.DocX.Paragraph.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Iterate through the paragraphs
    ///     foreach (Paragraph p in document.Paragraphs)
    ///     {
    ///         // Remove all but the first 2 characters from this Paragraph.
    ///         p.RemoveText(2, false);
    ///     }
    ///        
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
    /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
    /// <param name="index">The position to begin deleting characters.</param>
    /// <param name="trackChanges">Track changes</param>
    public void RemoveText( int index, bool trackChanges = false )
    {
      this.RemoveText( index, Text.Length - index, trackChanges );
    }

    /// <summary>
    /// Replaces all occurrences of a specified System.String in this instance, with another specified System.String.
    /// </summary>
    /// <example>
    /// <code>
    /// // Load a document using a relative filename.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // The formatting to match.
    ///     Formatting matchFormatting = new Formatting();
    ///     matchFormatting.Size = 10;
    ///     matchFormatting.Italic = true;
    ///     matchFormatting.FontFamily = new FontFamily("Times New Roman");
    ///
    ///     // The formatting to apply to the inserted text.
    ///     Formatting newFormatting = new Formatting();
    ///     newFormatting.Size = 22;
    ///     newFormatting.UnderlineStyle = UnderlineStyle.dotted;
    ///     newFormatting.Bold = true;
    ///
    ///     // Iterate through the paragraphs in this document.
    ///     foreach (Paragraph p in document.Paragraphs)
    ///     {
    ///         /* 
    ///          * Replace all instances of the string "wrong" with the string "right" and ignore case.
    ///          * Each inserted instance of "wrong" should use the Formatting newFormatting.
    ///          * Only replace an instance of "wrong" if it is Size 10, Italic and Times New Roman.
    ///          * SubsetMatch means that the formatting must contain all elements of the match formatting,
    ///          * but it can also contain additional formatting for example Color, UnderlineStyle, etc.
    ///          * ExactMatch means it must not contain additional formatting.
    ///          */
    ///         p.ReplaceText("wrong", "right", false, RegexOptions.IgnoreCase, newFormatting, matchFormatting, MatchFormattingOptions.SubsetMatch);
    ///     }
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <seealso cref="Paragraph.RemoveText(int, int, bool, bool)"/>
    /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
    /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
    /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
    /// <param name="newValue">A System.String to replace all occurrences of oldValue.</param>
    /// <param name="searchValue">A System.String to be replaced.</param>
    /// <param name="options">A bitwise OR combination of RegexOption enumeration options.</param>
    /// <param name="trackChanges">Track changes</param>
    /// <param name="newFormatting">The formatting to apply to the text being inserted.</param>
    /// <param name="matchFormatting">The formatting that the text must match in order to be replaced.</param>
    /// <param name="fo">How should formatting be matched?</param>
    /// <param name="escapeRegEx">True if the oldValue needs to be escaped, otherwise false. If it represents a valid RegEx pattern this should be false.</param>
    /// <param name="useRegExSubstitutions">True if RegEx-like replace should be performed, i.e. if newValue contains RegEx substitutions. Does not perform named-group substitutions (only numbered groups).</param>
    /// <param name="removeEmptyParagraph">Remove empty paragraph</param>
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
            var rPr = run.Xml.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );

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
            var rPr = run.Xml.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );
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

          // Replace text when formatting matches.
          if( formattingMatch )
          {
            var newValue = regexMatchHandler.Invoke( match.Groups[ 1 ].Value );
            this.InsertText( match.Index + match.Value.Length, newValue, trackChanges, newFormatting );
            this.RemoveText( match.Index, match.Value.Length, trackChanges, removeEmptyParagraph );
          }
        }
      }
    }

    /// <summary>
    /// Find all instances of a string in this paragraph and return their indexes in a List.
    /// </summary>
    /// <param name="str">The string to find</param>
    /// <returns>A list of indexes.</returns>
    /// <example>
    /// Find all instances of Hello in this document and insert 'don't' in frount of them.
    /// <code>
    /// // Load a document
    /// using (DocX document = DocX.Load(@"Test.docx"))
    /// {
    ///     // Loop through the paragraphs in this document.
    ///     foreach(Paragraph p in document.Paragraphs)
    ///     {
    ///         // Find all instances of 'go' in this paragraph.
    ///         <![CDATA[ List<int> ]]> gos = document.FindAll("go");
    ///
    ///         /* 
    ///          * Insert 'don't' in frount of every instance of 'go' in this document to produce 'don't go'.
    ///          * An important trick here is to do the inserting in reverse document order. If you inserted 
    ///          * in document order, every insert would shift the index of the remaining matches.
    ///          */
    ///         gos.Reverse();
    ///         foreach (int index in gos)
    ///         {
    ///             p.InsertText(index, "don't ", false);
    ///         }
    ///     }
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public List<int> FindAll( string str )
    {
      return this.FindAll( str, RegexOptions.None );
    }

    /// <summary>
    /// Find all instances of a string in this paragraph and return their indexes in a List.
    /// </summary>
    /// <param name="str">The string to find</param>
    /// <param name="options">The options to use when finding a string match.</param>
    /// <returns>A list of indexes.</returns>
    /// <example>
    /// Find all instances of Hello in this document and insert 'don't' in frount of them.
    /// <code>
    /// // Load a document
    /// using (DocX document = DocX.Load(@"Test.docx"))
    /// {
    ///     // Loop through the paragraphs in this document.
    ///     foreach(Paragraph p in document.Paragraphs)
    ///     {
    ///         // Find all instances of 'go' in this paragraph (Ignore case).
    ///         <![CDATA[ List<int> ]]>  gos = document.FindAll("go", RegexOptions.IgnoreCase);
    ///
    ///         /* 
    ///          * Insert 'don't' in frount of every instance of 'go' in this document to produce 'don't go'.
    ///          * An important trick here is to do the inserting in reverse document order. If you inserted 
    ///          * in document order, every insert would shift the index of the remaining matches.
    ///          */
    ///         gos.Reverse();
    ///         foreach (int index in gos)
    ///         {
    ///             p.InsertText(index, "don't ", false);
    ///         }
    ///     }
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
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

    /// <summary>
    ///  Find all unique instances of the given Regex Pattern
    /// </summary>
    /// <param name="str"></param>
    /// <param name="options"></param>
    /// <returns></returns>
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

    /// <summary>
    /// Insert a PageNumber place holder into a Paragraph.
    /// This place holder should only be inserted into a Header or Footer Paragraph.
    /// Word will not automatically update this field if it is inserted into a document level Paragraph.
    /// </summary>
    /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    /// <param name="index">The text index to insert this PageNumber place holder at.</param>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add Headers to the document.
    ///     document.AddHeaders();
    ///
    ///     // Get the default Header.
    ///     Header header = document.Headers.odd;
    ///
    ///     // Insert a Paragraph into the Header.
    ///     Paragraph p0 = header.InsertParagraph("Page ( of )");
    ///
    ///     // Insert place holders for PageNumber and PageCount into the Header.
    ///     // Word will replace these with the correct value for each Page.
    ///     p0.InsertPageNumber(PageNumberFormat.normal, 6);
    ///     p0.InsertPageCount(PageNumberFormat.normal, 11);
    ///
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AppendPageCount"/>
    /// <seealso cref="AppendPageNumber"/>
    /// <seealso cref="InsertPageCount"/>
    public void InsertPageNumber( PageNumberFormat pnf, int index = 0 )
    {
      var fldSimple = new XElement( XName.Get( "fldSimple", DocX.w.NamespaceName ) );

      if( pnf == PageNumberFormat.normal )
      {
        fldSimple.Add( new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), @" PAGE   \* MERGEFORMAT " ) );
      }
      else
      {
        fldSimple.Add( new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), @" PAGE  \* ROMAN  \* MERGEFORMAT " ) );
      }

      var content = XElement.Parse
      (
       @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof /> 
                   </w:rPr>
                   <w:t>1</w:t> 
               </w:r>"
      );

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
    }

    /// <summary>
    /// Append a PageNumber place holder onto the end of a Paragraph.
    /// </summary>
    /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add Headers to the document.
    ///     document.AddHeaders();
    ///
    ///     // Get the default Header.
    ///     Header header = document.Headers.odd;
    ///
    ///     // Insert a Paragraph into the Header.
    ///     Paragraph p0 = header.InsertParagraph();
    ///
    ///     // Appemd place holders for PageNumber and PageCount into the Header.
    ///     // Word will replace these with the correct value for each Page.
    ///     p0.Append("Page (");
    ///     p0.AppendPageNumber(PageNumberFormat.normal);
    ///     p0.Append(" of ");
    ///     p0.AppendPageCount(PageNumberFormat.normal);
    ///     p0.Append(")");
    /// 
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AppendPageCount"/>
    /// <seealso cref="InsertPageNumber"/>
    /// <seealso cref="InsertPageCount"/>
    public void AppendPageNumber( PageNumberFormat pnf )
    {
      XElement fldSimple = new XElement( XName.Get( "fldSimple", DocX.w.NamespaceName ) );

      if( pnf == PageNumberFormat.normal )
        fldSimple.Add( new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), @" PAGE   \* MERGEFORMAT " ) );
      else
        fldSimple.Add( new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), @" PAGE  \* ROMAN  \* MERGEFORMAT " ) );

      XElement content = XElement.Parse
      (
       @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof /> 
                   </w:rPr>
                   <w:t>1</w:t> 
               </w:r>"
      );

      fldSimple.Add( content );
      Xml.Add( fldSimple );
    }

    /// <summary>
    /// Insert a PageCount place holder into a Paragraph.
    /// This place holder should only be inserted into a Header or Footer Paragraph.
    /// Word will not automatically update this field if it is inserted into a document level Paragraph.
    /// </summary>
    /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    /// <param name="index">The text index to insert this PageCount place holder at.</param>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add Headers to the document.
    ///     document.AddHeaders();
    ///
    ///     // Get the default Header.
    ///     Header header = document.Headers.odd;
    ///
    ///     // Insert a Paragraph into the Header.
    ///     Paragraph p0 = header.InsertParagraph("Page ( of )");
    ///
    ///     // Insert place holders for PageNumber and PageCount into the Header.
    ///     // Word will replace these with the correct value for each Page.
    ///     p0.InsertPageNumber(PageNumberFormat.normal, 6);
    ///     p0.InsertPageCount(PageNumberFormat.normal, 11);
    ///
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AppendPageCount"/>
    /// <seealso cref="AppendPageNumber"/>
    /// <seealso cref="InsertPageNumber"/>
    public void InsertPageCount( PageNumberFormat pnf, int index = 0 )
    {
      XElement fldSimple = new XElement( XName.Get( "fldSimple", DocX.w.NamespaceName ) );

      if( pnf == PageNumberFormat.normal )
        fldSimple.Add( new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), @" NUMPAGES   \* MERGEFORMAT " ) );
      else
        fldSimple.Add( new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), @" NUMPAGES  \* ROMAN  \* MERGEFORMAT " ) );

      XElement content = XElement.Parse
      (
       @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof /> 
                   </w:rPr>
                   <w:t>1</w:t> 
               </w:r>"
      );

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
    }

    /// <summary>
    /// Append a PageCount place holder onto the end of a Paragraph.
    /// </summary>
    /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add Headers to the document.
    ///     document.AddHeaders();
    ///
    ///     // Get the default Header.
    ///     Header header = document.Headers.odd;
    ///
    ///     // Insert a Paragraph into the Header.
    ///     Paragraph p0 = header.InsertParagraph();
    ///
    ///     // Appemd place holders for PageNumber and PageCount into the Header.
    ///     // Word will replace these with the correct value for each Page.
    ///     p0.Append("Page (");
    ///     p0.AppendPageNumber(PageNumberFormat.normal);
    ///     p0.Append(" of ");
    ///     p0.AppendPageCount(PageNumberFormat.normal);
    ///     p0.Append(")");
    /// 
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AppendPageNumber"/>
    /// <seealso cref="InsertPageNumber"/>
    /// <seealso cref="InsertPageCount"/>
    public void AppendPageCount( PageNumberFormat pnf )
    {
      XElement fldSimple = new XElement( XName.Get( "fldSimple", DocX.w.NamespaceName ) );

      if( pnf == PageNumberFormat.normal )
        fldSimple.Add( new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), @" NUMPAGES   \* MERGEFORMAT " ) );
      else
        fldSimple.Add( new XAttribute( XName.Get( "instr", DocX.w.NamespaceName ), @" NUMPAGES  \* ROMAN  \* MERGEFORMAT " ) );

      XElement content = XElement.Parse
      (
       @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof /> 
                   </w:rPr>
                   <w:t>1</w:t> 
               </w:r>"
      );

      fldSimple.Add( content );
      Xml.Add( fldSimple );
    }

    /// <summary>
    /// Set the Line spacing for this paragraph manually.
    /// </summary>
    /// <param name="spacingType">The type of spacing to be set, can be either Before, After or Line (Standard line spacing).</param>
    /// <param name="spacingFloat">A float value of the amount of spacing. Equals the value that will be set in Word using the "Line and Paragraph spacing" button.</param>
    public void SetLineSpacing( LineSpacingType spacingType, float spacingFloat )
    {
      var pPr = this.GetOrCreate_pPr();
      var spacingXName = XName.Get( "spacing", DocX.w.NamespaceName );
      var spacing = pPr.Element( spacingXName );
      if( spacing == null )
      {
        pPr.Add( new XElement( spacingXName ) );
        spacing = pPr.Element( spacingXName );
      }

      var spacingTypeAttribute = ( spacingType == LineSpacingType.Before )
                                 ? "before"
                                 : ( spacingType == LineSpacingType.After ) ? "after" : "line";
      spacing.SetAttributeValue( XName.Get( spacingTypeAttribute, DocX.w.NamespaceName ), ( int )( spacingFloat * 240f ) );
    }

    /// <summary>
    /// Set the Line spacing for this paragraph using the Auto value.
    /// </summary>
    /// <param name="spacingTypeAuto">The type of spacing to be set automatically. Using Auto will set both Before and After. None will remove any line spacing.</param>
    public void SetLineSpacing( LineSpacingTypeAuto spacingTypeAuto )
    {
      var pPr = this.GetOrCreate_pPr();
      var spacingXName = XName.Get( "spacing", DocX.w.NamespaceName );
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

        spacing.SetAttributeValue( XName.Get( spacingTypeAttribute, DocX.w.NamespaceName ), spacingValue );
        spacing.SetAttributeValue( XName.Get( autoSpacingTypeAttribute, DocX.w.NamespaceName ), 1 );

        if( spacingTypeAuto == LineSpacingTypeAuto.Auto )
        {
          spacing.SetAttributeValue( XName.Get( "after", DocX.w.NamespaceName ), spacingValue );
          spacing.SetAttributeValue( XName.Get( "afterAutospacing", DocX.w.NamespaceName ), 1 );
        }
      }
    }

    public Paragraph AppendBookmark( string bookmarkName )
    {
      XElement wBookmarkStart = new XElement(
          XName.Get( "bookmarkStart", DocX.w.NamespaceName ),
          new XAttribute( XName.Get( "id", DocX.w.NamespaceName ), 0 ),
          new XAttribute( XName.Get( "name", DocX.w.NamespaceName ), bookmarkName ) );
      Xml.Add( wBookmarkStart );

      XElement wBookmarkEnd = new XElement(
          XName.Get( "bookmarkEnd", DocX.w.NamespaceName ),
          new XAttribute( XName.Get( "id", DocX.w.NamespaceName ), 0 ),
          new XAttribute( XName.Get( "name", DocX.w.NamespaceName ), bookmarkName ) );
      Xml.Add( wBookmarkEnd );

      return this;
    }

    public bool ValidateBookmark( string bookmarkName )
    {
      return GetBookmarks().Any( b => b.Name.Equals( bookmarkName ) );
    }

    public IEnumerable<Bookmark> GetBookmarks()
    {
      return Xml.Descendants( XName.Get( "bookmarkStart", DocX.w.NamespaceName ) )
          .Select( x => x.Attribute( XName.Get( "name", DocX.w.NamespaceName ) ) )
          .Select( x => new Bookmark
          {
            Name = x.Value,
            Paragraph = this
          } );
    }

    public void InsertAtBookmark( string toInsert, string bookmarkName )
    {
      var bookmark = Xml.Descendants( XName.Get( "bookmarkStart", DocX.w.NamespaceName ) )
                          .Where( x => x.Attribute( XName.Get( "name", DocX.w.NamespaceName ) ).Value == bookmarkName ).SingleOrDefault();
      if( bookmark != null )
      {
        var run = HelperFunctions.FormatInput( toInsert, null );
        bookmark.AddBeforeSelf( run );
        _runs = Xml.Elements( XName.Get( "r", DocX.w.NamespaceName ) ).ToList();
        HelperFunctions.RenumberIDs( Document );
      }
    }

    /// <summary>
    /// Paragraph that will be kept on the same page as the next paragraph.
    /// </summary>
    /// <param name="keepWithNextParagraph"></param>
    /// <returns></returns>
    public Paragraph KeepWithNextParagraph( bool keepWithNextParagraph = true )
    {
      var pPr = GetOrCreate_pPr();
      var keepNextElement = pPr.Element( XName.Get( "keepNext", DocX.w.NamespaceName ) );

      if( keepNextElement == null && keepWithNextParagraph )
      {
        pPr.Add( new XElement( XName.Get( "keepNext", DocX.w.NamespaceName ) ) );
      }

      if( !keepWithNextParagraph && keepNextElement != null )
      {
        keepNextElement.Remove();
      }

      return this;
    }

    /// <summary>
    /// Paragraph with lines that will stay together on the same page.
    /// </summary>
    /// <param name="keepLinesTogether"></param>
    /// <returns></returns>
    public Paragraph KeepLinesTogether( bool keepLinesTogether = true )
    {
      var pPr = GetOrCreate_pPr();
      var keepLinesElement = pPr.Element( XName.Get( "keepLines", DocX.w.NamespaceName ) );

      if( keepLinesElement == null && keepLinesTogether )
      {
        pPr.Add( new XElement( XName.Get( "keepLines", DocX.w.NamespaceName ) ) );
      }

      if( !keepLinesTogether )
      {
        keepLinesElement?.Remove();
      }

      return this;
    }

    public void ReplaceAtBookmark( string text, string bookmarkName )
    {
      var bookmarkStart = this.Xml.Descendants( XName.Get( "bookmarkStart", DocX.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "name", DocX.w.NamespaceName ) ).Value == bookmarkName )
                                  .SingleOrDefault();
      if( bookmarkStart == null )
        return;

      var nextNode = bookmarkStart.NextNode;
      XElement nextXElement = null;
      while( nextNode != null )
      {
        nextXElement = nextNode as XElement;
        if( ( nextXElement != null ) &&
            ( nextXElement.Name.NamespaceName == DocX.w.NamespaceName ) &&
            ( ( nextXElement.Name.LocalName == "r" ) || ( nextXElement.Name.LocalName == "bookmarkEnd" ) ) )
          break;

        nextNode = nextNode.NextNode;
      }

      if( nextXElement == null )
        return;

      if( nextXElement.Name.LocalName.Equals( "bookmarkEnd" ) )
      {
        this.ReplaceAtBookmark_Core( text, bookmarkStart );
        return;
      }

      var tXElement = nextXElement.Elements( XName.Get( "t", DocX.w.NamespaceName ) ).FirstOrDefault();
      if( tXElement == null )
      {
        this.ReplaceAtBookmark_Core( text, bookmarkStart );
        return;
      }

      tXElement.Value = text;
    }

    public void InsertHorizontalLine( HorizontalBorderPosition position = HorizontalBorderPosition.bottom, string lineType = "single", int size = 6, int space = 1, string color = "auto" )
    {
      var pBrXName = XName.Get( "pBdr", DocX.w.NamespaceName );
      var borderPositionXName = ( position == HorizontalBorderPosition.bottom) ? XName.Get( "bottom", DocX.w.NamespaceName ) : XName.Get( "top", DocX.w.NamespaceName );

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
        border.SetAttributeValue( XName.Get( "val", DocX.w.NamespaceName ), lineType );
        border.SetAttributeValue( XName.Get( "sz", DocX.w.NamespaceName ), size.ToString() );
        border.SetAttributeValue( XName.Get( "space", DocX.w.NamespaceName ), space.ToString() );
        border.SetAttributeValue( XName.Get( "color", DocX.w.NamespaceName ), color );
      }
    }

    #endregion

    #region Internal Methods

    /// <summary>
    /// If the pPr element doesent exist it is created, either way it is returned by this function.
    /// </summary>
    /// <returns>The pPr element for this Paragraph.</returns>
    internal XElement GetOrCreate_pPr()
    {
      // Get the element.
      var pPr = Xml.Element( XName.Get( "pPr", DocX.w.NamespaceName ) );

      // If it dosen't exist, create it.
      if( pPr == null )
      {
        Xml.AddFirst( new XElement( XName.Get( "pPr", DocX.w.NamespaceName ) ) );
        pPr = Xml.Element( XName.Get( "pPr", DocX.w.NamespaceName ) );
      }

      // Return the pPr element for this Paragraph.
      return pPr;
    }

    /// <summary>
    /// If the ind element doesent exist it is created, either way it is returned by this function.
    /// </summary>
    /// <returns>The ind element for this Paragraphs pPr.</returns>
    internal XElement GetOrCreate_pPr_ind()
    {
      // Get the element.
      XElement pPr = GetOrCreate_pPr();
      XElement ind = pPr.Element( XName.Get( "ind", DocX.w.NamespaceName ) );

      // If it dosen't exist, create it.
      if( ind == null )
      {
        pPr.Add( new XElement( XName.Get( "ind", DocX.w.NamespaceName ) ) );
        ind = pPr.Element( XName.Get( "ind", DocX.w.NamespaceName ) );
      }

      // Return the pPr element for this Paragraph.
      return ind;
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

    /// <summary>
    /// Create a new Picture.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="id">A unique id that identifies an Image embedded in this document.</param>
    /// <param name="name">The name of this Picture.</param>
    /// <param name="descr">The description of this Picture.</param>
    /// <param name="width">The width of this Picture.</param>
    /// <param name="height">The height of this Picture.</param>
    static internal Picture CreatePicture( DocX document, string id, string name, string descr, int width, int height )
    {
      var part = document._package.GetPart( document.PackagePart.GetRelationship( id ).TargetUri );

      long newDocPrId = document.GetNextFreeDocPrId();
      int cx, cy;

      using( PackagePartStream packagePartStream = new PackagePartStream( part.GetStream() ) )
      {
        using( System.Drawing.Image img = System.Drawing.Image.FromStream( packagePartStream, useEmbeddedColorManagement: false, validateImageData: false ) )
        {
          cx = img.Width * 9526;
          cy = img.Height * 9526;
        }
      }

      var xml = XElement.Parse
      ( string.Format( @"
        <w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
            <w:drawing xmlns = ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                    <wp:extent cx=""{0}"" cy=""{1}"" />
                    <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0"" />
                    <wp:docPr id=""{5}"" name=""{3}"" descr=""{4}"" />
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
        ", cx, cy, id, name, descr, newDocPrId.ToString() ) );

      var picture = new Picture( document, xml, new Image( document, document.PackagePart.GetRelationship( id ) ) );
      if( width > -1 )
      {
        picture.Width = width;
      }
      if( height > -1 )
      {
        picture.Height = height;
      }
      return picture;
    }

    /// <summary>
    /// Creates an Edit either a ins or a del with the specified content and date
    /// </summary>
    /// <param name="t">The type of this edit (ins or del)</param>
    /// <param name="edit_time">The time stamp to use for this edit</param>
    /// <param name="content">The initial content of this edit</param>
    /// <returns></returns>
    internal static XElement CreateEdit( EditType t, DateTime edit_time, object content )
    {
      if( t == EditType.del )
      {
        foreach( object o in ( IEnumerable<XElement> )content )
        {
          if( o is XElement )
          {
            XElement e = ( o as XElement );
            IEnumerable<XElement> ts = e.DescendantsAndSelf( XName.Get( "t", DocX.w.NamespaceName ) );

            for( int i = 0; i < ts.Count(); i++ )
            {
              XElement text = ts.ElementAt( i );
              text.ReplaceWith( new XElement( DocX.w + "delText", text.Attributes(), text.Value ) );
            }
          }
        }
      }

      // Check the author in a Try/Catch 
      // (for the cases where we do not have the rights to access that information)
      string author = "";
      try
      {
        author = WindowsIdentity.GetCurrent().Name;
      }
      catch( Exception )
      {
        // do nothing
      }

      if( author.Trim() == "" )
      {
        return
        (
            new XElement( DocX.w + t.ToString(),
                new XAttribute( DocX.w + "id", 0 ),
                new XAttribute( DocX.w + "date", edit_time ),
            content )
        );
      }

      return
      (
          new XElement( DocX.w + t.ToString(),
              new XAttribute( DocX.w + "id", 0 ),
              new XAttribute( DocX.w + "author", author ),
              new XAttribute( DocX.w + "date", edit_time ),
          content )
      );
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

    internal void GetFirstRunEffectedByEditRecursive( XElement Xml, int index, ref int count, ref Run theOne, EditType type )
    {
      count += HelperFunctions.GetSize( Xml );

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

      if( Xml.HasElements )
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

    /// <!-- 
    /// Bug found and fixed by krugs525 on August 12 2009.
    /// Use TFS compare to see exact code change.
    /// -->
    static internal int GetElementTextLength( XElement run )
    {
      int count = 0;

      if( run == null )
        return count;

      foreach( var d in run.Descendants() )
      {
        switch( d.Name.LocalName )
        {
          case "tab":
            if( d.Parent.Name.LocalName != "tabs" )
              goto case "br";
            break;
          case "br":
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

    internal string GetOrGenerateRel( Picture p )
    {
      string image_uri_string = p._img._pr.TargetUri.OriginalString;

      // Search for a relationship with a TargetUri that points at this Image.
      string id = null;
      foreach( var r in this.PackagePart.GetRelationshipsByType( DocX.RelationshipImage ) )
      {
        if( string.Equals( r.TargetUri.OriginalString, image_uri_string, StringComparison.Ordinal ) )
        {
          id = r.Id;
          break;
        }
      }

      // If such a relation doesn't exist, create one.
      if( id == null )
      {
        // Check to see if a relationship for this Picture exists and create it if not.
        var pr = this.PackagePart.CreateRelationship( p._img._pr.TargetUri, TargetMode.Internal, DocX.RelationshipImage );
        id = pr.Id;
      }
      return id;
    }

    internal string GetOrGenerateRel( Hyperlink h )
    {
      string image_uri_string = (h.Uri != null) ?  h.Uri.OriginalString : null;

      // Search for a relationship with a TargetUri that points at this Image.
      var Id =
      (
          from r in this.PackagePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" )
          where r.TargetUri.OriginalString == image_uri_string
          select r.Id
      ).SingleOrDefault();

      // If such a relation dosen't exist, create one.
      if( (Id == null) && ( h.Uri != null) )
      {
        // Check to see if a relationship for this Picture exists and create it if not.
        var pr = this.PackagePart.CreateRelationship( h.Uri, TargetMode.External, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" );
        Id = pr.Id;
      }
      return Id;
    }

    internal void ApplyTextFormattingProperty( XName textFormatPropName, string value, object content )
    {
      XElement rPr = null;

      if( _runs.Count == 0 )
      {
        var pPr = this.Xml.Element( XName.Get( "pPr", DocX.w.NamespaceName ) );
        if( pPr == null )
        {
          this.Xml.AddFirst( new XElement( XName.Get( "pPr", DocX.w.NamespaceName ) ) );
          pPr = this.Xml.Element( XName.Get( "pPr", DocX.w.NamespaceName ) );
        }

        rPr = pPr.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );
        if( rPr == null )
        {
          pPr.AddFirst( new XElement( XName.Get( "rPr", DocX.w.NamespaceName ) ) );
          rPr = pPr.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );
        }

        rPr.SetElementValue( textFormatPropName, value );

        var lastElement = rPr.Elements( textFormatPropName ).Last();
        // Check if the content is an attribute
        if( content as XAttribute != null )
        {
          // Add or Update the attribute to the last element
          if( lastElement.Attribute( ( ( XAttribute )( content ) ).Name ) == null )
          {
            lastElement.Add( content );
          }
          else
          {
            lastElement.Attribute( ( ( XAttribute )( content ) ).Name ).Value = ( ( XAttribute )( content ) ).Value;
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
        rPr = run.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );
        if( rPr == null )
        {
          run.AddFirst( new XElement( XName.Get( "rPr", DocX.w.NamespaceName ) ) );
          rPr = run.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );
        }

        rPr.SetElementValue( textFormatPropName, value );
        var last = rPr.Elements( textFormatPropName ).Last();

        if( isFontPropertiesList )
        {
          foreach( object property in fontProperties )
          {
            if( last.Attribute( ( ( XAttribute )( property ) ).Name ) == null )
            {
              last.Add( property );
            }
            else
            {
              last.Attribute( ( ( XAttribute )( property ) ).Name ).Value = ( ( XAttribute )( property ) ).Value;
            }
          }
        }


        if( content as XAttribute != null )//If content is an attribute
        {
          if( last.Attribute( ( ( XAttribute )( content ) ).Name ) == null )
          {
            last.Add( content ); //Add this attribute if element doesn't have it
          }
          else
          {
            last.Attribute( ( ( XAttribute )( content ) ).Name ).Value = ( ( XAttribute )( content ) ).Value; //Apply value only if element already has it
          }
        }
        else
        {
          //IMPORTANT
          //But what to do if it is not?
        }
      }
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
      docProperties =
      (
          from xml in Xml.Descendants( XName.Get( "fldSimple", DocX.w.NamespaceName ) )
          select new DocProperty( Document, xml )
      ).ToList();
    }

    private XElement GetParagraphNumberProperties()
    {
      var numPrNode = Xml.Descendants().FirstOrDefault( el => el.Name.LocalName == "numPr" );
      return numPrNode;
    }

    private List<Picture> GetPictures( string localName, string localNameEquals, string attributeName )
    {
      var pictures =
       (
           from p in Xml.Descendants()
           where ( p.Name.LocalName == localName )
           let id =
           (
               from e in p.Descendants()
               where e.Name.LocalName.Equals( localNameEquals )
               select e.Attribute( XName.Get( attributeName, "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) ).Value
           ).SingleOrDefault()
           where id != null
           let img = new Image( this.Document, this.PackagePart.GetRelationship( id ) )
           select new Picture( this.Document, p, img )
       ).ToList();

      return pictures;
    }

    private void ReplaceAtBookmark_Core( string text, XElement bookmark )
    {
      var xElementList = HelperFunctions.FormatInput( text, null );
      bookmark.AddAfterSelf( xElementList );

      _runs = this.Xml.Elements( XName.Get( "r", DocX.w.NamespaceName ) ).ToList();

      HelperFunctions.RenumberIDs( this.Document );
    }

    #endregion

  }

  public class Run : DocXElement
  {
    #region Private Members

    // A lookup for the text elements in this paragraph
    private Dictionary<int, Text> textLookup = new Dictionary<int, Text>();

    private int startIndex;
    private int endIndex;
    private string text;

    #endregion

    #region Public Properties

    /// <summary>
    /// Gets the start index of this Text (text length before this text)
    /// </summary>
    public int StartIndex
    {
      get
      {
        return startIndex;
      }
    }

    /// <summary>
    /// Gets the end index of this Text (text length before this text + this texts length)
    /// </summary>
    public int EndIndex
    {
      get
      {
        return endIndex;
      }
    }

    #endregion

    #region Internal Properties

    /// <summary>
    /// The text value of this text element
    /// </summary>
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

    internal Run( DocX document, XElement xml, int startIndex )
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
              textLookup.Add( start + 1, new Text( Document, te, start ) );
              text += "\n";
              start++;
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

    #region Iternal Methods

    static internal XElement[] SplitRun( Run r, int index, EditType type = EditType.ins )
    {
      index = index - r.StartIndex;

      Text t = r.GetFirstTextEffectedByEdit( index, type );
      XElement[] splitText = Text.SplitText( t, index );

      XElement splitLeft = new XElement( r.Xml.Name, r.Xml.Attributes(), r.Xml.Element( XName.Get( "rPr", DocX.w.NamespaceName ) ), t.Xml.ElementsBeforeSelf().Where( n => n.Name.LocalName != "rPr" ), splitText[ 0 ] );
      if( Paragraph.GetElementTextLength( splitLeft ) == 0 )
        splitLeft = null;

      XElement splitRight = new XElement( r.Xml.Name, r.Xml.Attributes(), r.Xml.Element( XName.Get( "rPr", DocX.w.NamespaceName ) ), splitText[ 1 ], t.Xml.ElementsAfterSelf().Where( n => n.Name.LocalName != "rPr" ) );
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

  internal class Text : DocXElement
  {
    #region Private Members

    private int startIndex;
    private int endIndex;
    private string text;

    #endregion

    #region Public Properties

    /// <summary>
    /// Gets the start index of this Text (text length before this text)
    /// </summary>
    public int StartIndex
    {
      get
      {
        return startIndex;
      }
    }

    /// <summary>
    /// Gets the end index of this Text (text length before this text + this texts length)
    /// </summary>
    public int EndIndex
    {
      get
      {
        return endIndex;
      }
    }

    /// <summary>
    /// The text value of this text element
    /// </summary>
    public string Value
    {
      get
      {
        return text;
      }
    }

    #endregion

    #region Constructors

    internal Text( DocX document, XElement xml, int startIndex )
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
            text = "\n";
            endIndex = startIndex + 1;
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

    /// <summary>
    /// If a text element or delText element, starts or ends with a space,
    /// it must have the attribute space, otherwise it must not have it.
    /// </summary>
    /// <param name="e">The (t or delText) element check</param>
    public static void PreserveSpace( XElement e )
    {
      // PreserveSpace should only be used on (t or delText) elements
      if( !e.Name.Equals( DocX.w + "t" ) && !e.Name.Equals( DocX.w + "delText" ) )
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
