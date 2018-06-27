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
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using System.Security.Cryptography;
using System.Drawing;
using System.Collections.ObjectModel;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Represents a document.
  /// </summary>
  public class DocX : Container, IDisposable
  {
    #region Namespaces
    static internal XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static internal XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    static internal XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    static internal XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    static internal XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    static internal XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    static internal XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    static internal XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    static internal XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    internal static XNamespace n = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    static internal XNamespace v = "urn:schemas-microsoft-com:vml";
    #endregion

    #region Private Members

    private Headers _headers;
    private Footers _footers;

    private float _pageSizeMultiplier = 20.0f;

    private readonly object nextFreeDocPrIdLock = new object();
    private long? nextFreeDocPrId;

    #endregion

    #region Internal Constants

    internal const string RelationshipImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    internal const string ContentTypeApplicationRelationShipXml = "application/vnd.openxmlformats-package.relationships+xml";

    #endregion

    #region Internal Members

    // Get the word\settings.xml part
    internal PackagePart _settingsPart;
    internal PackagePart _endnotesPart;
    internal PackagePart _footnotesPart;
    internal PackagePart _stylesPart;
    internal PackagePart _stylesWithEffectsPart;
    internal PackagePart _numberingPart;
    internal PackagePart _fontTablePart;

    #region Internal variables defined foreach DocX object
    // Object representation of the .docx
    internal Package _package;

    // The mainDocument is loaded into a XDocument object for easy querying and editing
    internal XDocument _mainDoc;
    internal XDocument _settings;
    internal XDocument _endnotes;
    internal XDocument _footnotes;
    internal XDocument _styles;
    internal XDocument _stylesWithEffects;
    internal XDocument _numbering;
    internal XDocument _fontTable;

    // A lookup for the Paragraphs in this document.
    internal Dictionary<int, Paragraph> _paragraphLookup = new Dictionary<int, Paragraph>();
    // Every document is stored in a MemoryStream, all edits made to a document are done in memory.
    internal MemoryStream _memoryStream;
    // The filename that this document was loaded from
    internal string _filename;
    // The stream that this document was loaded from
    internal Stream _stream;
    #endregion

    #endregion

    #region Public Properties

    /// <summary>
		/// Top margin in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
    public float MarginTop
    {
      get
      {
        return getMarginAttribute( XName.Get( "top", w.NamespaceName ) );
      }

      set
      {
        setMarginAttribute( XName.Get( "top", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Bottom margin in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginBottom
    {
      get
      {
        return getMarginAttribute( XName.Get( "bottom", w.NamespaceName ) );
      }

      set
      {
        setMarginAttribute( XName.Get( "bottom", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Left margin in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginLeft
    {
      get
      {
        return getMarginAttribute( XName.Get( "left", w.NamespaceName ) );
      }

      set
      {
        setMarginAttribute( XName.Get( "left", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Right margin in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginRight
    {
      get
      {
        return getMarginAttribute( XName.Get( "right", w.NamespaceName ) );
      }

      set
      {
        setMarginAttribute( XName.Get( "right", w.NamespaceName ), value );
      }
    }

    /// <summary>
    /// Header margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginHeader
    {
      get
      {
        return getMarginAttribute(XName.Get("header", w.NamespaceName));
      }
      set
      {
        setMarginAttribute(XName.Get("header", w.NamespaceName), value);
      }
    }

    /// <summary>
    /// Footer margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float MarginFooter
    {
      get
      {
        return getMarginAttribute(XName.Get("footer", w.NamespaceName));
      }
      set
      {
        setMarginAttribute(XName.Get("footer", w.NamespaceName), value);
      }
    }

    public bool MirrorMargins
    {
      get
      {
        return getMirrorMargins(XName.Get("mirrorMargins", DocX.w.NamespaceName));
      }
      set
      {
        setMirrorMargins(XName.Get("mirrorMargins", DocX.w.NamespaceName), value);
      }
    }

    /// <summary>
    /// Page width in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float PageWidth
    {
      get
      {
        var body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
        var sectPr = body.Element( XName.Get( "sectPr", w.NamespaceName ) );
        var pgSz = sectPr?.Element( XName.Get( "pgSz", w.NamespaceName ) );

        if( pgSz != null )
        {
          var w = pgSz.Attribute( XName.Get( "w", DocX.w.NamespaceName ) );
          if( w != null )
          {
            float f;
            if( float.TryParse( w.Value, out f ) )
              return ( int )( f / _pageSizeMultiplier );
          }
        }

        return ( 12240.0f / _pageSizeMultiplier );
      }

      set
      {
        var body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
        var sectPr = body.Element( XName.Get( "sectPr", w.NamespaceName ) );
        var pgSz = sectPr.Element( XName.Get( "pgSz", w.NamespaceName ) );
        pgSz?.SetAttributeValue( XName.Get( "w", w.NamespaceName ), value * Convert.ToInt32( _pageSizeMultiplier ) );
      }
    }

    /// <summary>
    /// Page height in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public float PageHeight
    {
      get
      {
        var body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
        var sectPr = body.Element( XName.Get( "sectPr", w.NamespaceName ) );
        if( sectPr != null )
        {
          var pgSz = sectPr.Element( XName.Get( "pgSz", w.NamespaceName ) );
          if( pgSz != null )
          {
            var w = pgSz.Attribute( XName.Get( "h", DocX.w.NamespaceName ) );
            if( w != null )
            {
              float f;
              if( float.TryParse( w.Value, out f ) )
                return ( int )( f / _pageSizeMultiplier );
            }
          }
        }

        return ( 15840.0f / _pageSizeMultiplier );
      }

      set
      {
        var body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
        if( body != null )
        {
          var sectPr = body.Element( XName.Get( "sectPr", w.NamespaceName ) );
          if( sectPr != null )
          {
            var pgSz = sectPr.Element( XName.Get( "pgSz", w.NamespaceName ) );
            if( pgSz != null )
            {
              pgSz.SetAttributeValue( XName.Get( "h", w.NamespaceName ), value * Convert.ToInt32( _pageSizeMultiplier ) ); 
            }
          }
        }
      }
    }
    /// <summary>
    /// Returns true if any editing restrictions are imposed on this document.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     if(document.isProtected)
    ///         Console.WriteLine("Protected");
    ///     else
    ///         Console.WriteLine("Not protected");
    ///         
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AddProtection(EditRestrictions)"/>
    /// <seealso cref="RemoveProtection"/>
    /// <seealso cref="GetProtectionType"/>
    public bool isProtected
    {
      get
      {
        return _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).Count() > 0;
      }
    }

    public PageLayout PageLayout
    {
      get
      {
        XElement sectPr = Xml.Element( XName.Get( "sectPr", w.NamespaceName ) );
        if( sectPr == null )
        {
          Xml.SetElementValue( XName.Get( "sectPr", w.NamespaceName ), string.Empty );
          sectPr = Xml.Element( XName.Get( "sectPr", w.NamespaceName ) );
        }

        return new PageLayout( this, sectPr );
      }
    }

    /// <summary>
    /// Returns a collection of Headers in this Document.
    /// A document typically contains three Headers.
    /// A default one (odd), one for the first page and one for even pages.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///    // Add header support to this document.
    ///    document.AddHeaders();
    ///
    ///    // Get a collection of all headers in this document.
    ///    Headers headers = document.Headers;
    ///
    ///    // The header used for the first page of this document.
    ///    Header first = headers.first;
    ///
    ///    // The header used for odd pages of this document.
    ///    Header odd = headers.odd;
    ///
    ///    // The header used for even pages of this document.
    ///    Header even = headers.even;
    /// }
    /// </code>
    /// </example>
    public Headers Headers
    {
      get
      {
        return _headers;
      }
    }

    /// <summary>
    /// Returns a collection of Footers in this Document.
    /// A document typically contains three Footers.
    /// A default one (odd), one for the first page and one for even pages.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///    // Add footer support to this document.
    ///    document.AddFooters();
    ///
    ///    // Get a collection of all footers in this document.
    ///    Footers footers = document.Footers;
    ///
    ///    // The footer used for the first page of this document.
    ///    Footer first = footers.first;
    ///
    ///    // The footer used for odd pages of this document.
    ///    Footer odd = footers.odd;
    ///
    ///    // The footer used for even pages of this document.
    ///    Footer even = footers.even;
    /// }
    /// </code>
    /// </example>
    public Footers Footers
    {
      get
      {
        return _footers;
      }
    }

    /// <summary>
    /// Should the Document use different Headers and Footers for odd and even pages?
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add header support to this document.
    ///     document.AddHeaders();
    ///
    ///     // Get a collection of all headers in this document.
    ///     Headers headers = document.Headers;
    ///
    ///     // The header used for odd pages of this document.
    ///     Header odd = headers.odd;
    ///
    ///     // The header used for even pages of this document.
    ///     Header even = headers.even;
    ///
    ///     // Force the document to use a different header for odd and even pages.
    ///     document.DifferentOddAndEvenPages = true;
    ///
    ///     // Content can be added to the Headers in the same manor that it would be added to the main document.
    ///     Paragraph p1 = odd.InsertParagraph();
    ///     p1.Append("This is the odd pages header.");
    ///     
    ///     Paragraph p2 = even.InsertParagraph();
    ///     p2.Append("This is the even pages header.");
    ///
    ///     // Save all changes to this document.
    ///     document.Save();    
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public bool DifferentOddAndEvenPages
    {
      get
      {
        XDocument settings;
        using( TextReader tr = new StreamReader( _settingsPart.GetStream() ) )
        {
          settings = XDocument.Load( tr );
        }

        var evenAndOddHeaders = settings.Root.Element( w + "evenAndOddHeaders" );

        return ( evenAndOddHeaders != null );
      }

      set
      {
        XDocument settings;
        using( TextReader tr = new StreamReader( _settingsPart.GetStream() ) )
        {
          settings = XDocument.Load( tr );
        }

        var evenAndOddHeaders = settings.Root.Element( w + "evenAndOddHeaders" );
        if( evenAndOddHeaders == null )
        {
          if( value )
          {
            settings.Root.AddFirst( new XElement( w + "evenAndOddHeaders" ) );
          }
        }
        else
        {
          if( !value )
          {
            evenAndOddHeaders.Remove();
          }
        }

        using( TextWriter tw = new StreamWriter( new PackagePartStream( _settingsPart.GetStream() ) ) )
        {
          settings.Save( tw );
        }
      }
    }

    /// <summary>
    /// Should the Document use an independent Header and Footer for the first page?
    /// </summary>
    /// <example>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add header support to this document.
    ///     document.AddHeaders();
    ///
    ///     // The header used for the first page of this document.
    ///     Header first = document.Headers.first;
    ///
    ///     // Force the document to use a different header for first page.
    ///     document.DifferentFirstPage = true;
    ///     
    ///     // Content can be added to the Headers in the same manor that it would be added to the main document.
    ///     Paragraph p = first.InsertParagraph();
    ///     p.Append("This is the first pages header.");
    ///
    ///     // Save all changes to this document.
    ///     document.Save();    
    /// }// Release this document from memory.
    /// </example>
    public bool DifferentFirstPage
    {
      get
      {
        var body = _mainDoc.Root.Element( w + "body" );
        var sectPr = body.Element( w + "sectPr" );
        var titlePg = sectPr?.Element( w + "titlePg" );
        return titlePg != null;
      }

      set
      {
        var body = _mainDoc.Root.Element( w + "body" );
        body.Add( new XElement( w + "sectPr", string.Empty ) );
        var sectPr = body.Element( w + "sectPr" );
        var titlePg = sectPr.Element( w + "titlePg" );
        if( titlePg == null )
        {
          if( value )
          {
            sectPr.Add( new XElement( w + "titlePg", string.Empty ) );
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

    /// <summary>
    /// Returns a list of Images in this document.
    /// </summary>
    /// <example>
    /// Get the unique Id of every Image in this document.
    /// <code>
    /// // Load a document.
    /// DocX document = DocX.Load(@"C:\Example\Test.docx");
    ///
    /// // Loop through each Image in this document.
    /// foreach (Xceed.Words.NET.Image i in document.Images)
    /// {
    ///     // Get the unique Id which identifies this Image.
    ///     string uniqueId = i.Id;
    /// }
    ///
    /// </code>
    /// </example>
    /// <seealso cref="AddImage(string)"/>
    /// <seealso cref="AddImage(Stream, string)"/>
    /// <seealso cref="Paragraph.Pictures"/>
    /// <seealso cref="Paragraph.InsertPicture"/>
    public List<Image> Images
    {
      get
      {
        var imageRelationships = this.PackagePart.GetRelationshipsByType( RelationshipImage );
        if( imageRelationships.Any() )
        {
          return
          (
              from i in imageRelationships
              select new Image( this, i )
          ).ToList();
        }

        return new List<Image>();
      }
    }

    /// <summary>
    /// Returns a list of custom properties in this document.
    /// </summary>
    /// <example>
    /// Method 1: Get the name, type and value of each CustomProperty in this document.
    /// <code>
    /// // Load Example.docx
    /// DocX document = DocX.Load(@"C:\Example\Test.docx");
    ///
    /// /*
    ///  * No two custom properties can have the same name,
    ///  * so a Dictionary is the perfect data structure to store them in.
    ///  * Each custom property can be accessed using its name.
    ///  */
    /// foreach (string name in document.CustomProperties.Keys)
    /// {
    ///     // Grab a custom property using its name.
    ///     CustomProperty cp = document.CustomProperties[name];
    ///
    ///     // Write this custom properties details to Console.
    ///     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
    /// }
    ///
    /// Console.WriteLine("Press any key...");
    ///
    /// // Wait for the user to press a key before closing the Console.
    /// Console.ReadKey();
    /// </code>
    /// </example>
    /// <example>
    /// Method 2: Get the name, type and value of each CustomProperty in this document.
    /// <code>
    /// // Load Example.docx
    /// DocX document = DocX.Load(@"C:\Example\Test.docx");
    /// 
    /// /*
    ///  * No two custom properties can have the same name,
    ///  * so a Dictionary is the perfect data structure to store them in.
    ///  * The values of this Dictionary are CustomProperties.
    ///  */
    /// foreach (CustomProperty cp in document.CustomProperties.Values)
    /// {
    ///     // Write this custom properties details to Console.
    ///     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
    /// }
    ///
    /// Console.WriteLine("Press any key...");
    ///
    /// // Wait for the user to press a key before closing the Console.
    /// Console.ReadKey();
    /// </code>
    /// </example>
    /// <seealso cref="AddCustomProperty"/>
    public Dictionary<string, CustomProperty> CustomProperties
    {
      get
      {
        if( _package.PartExists( new Uri( "/docProps/custom.xml", UriKind.Relative ) ) )
        {
          PackagePart docProps_custom = _package.GetPart( new Uri( "/docProps/custom.xml", UriKind.Relative ) );
          XDocument customPropDoc;
          using( TextReader tr = new StreamReader( docProps_custom.GetStream( FileMode.Open, FileAccess.Read ) ) )
            customPropDoc = XDocument.Load( tr, LoadOptions.PreserveWhitespace );

          // Get all of the custom properties in this document
          return
          (
              from p in customPropDoc.Descendants( XName.Get( "property", customPropertiesSchema.NamespaceName ) )
              let Name = p.Attribute( XName.Get( "name" ) ).Value
              let Type = p.Descendants().Single().Name.LocalName
              let Value = p.Descendants().Single().Value
              select new CustomProperty( Name, Type, Value )
          ).ToDictionary( p => p.Name, StringComparer.CurrentCultureIgnoreCase );
        }

        return new Dictionary<string, CustomProperty>();
      }
    }

    ///<summary>
    /// Returns the list of document core properties with corresponding values.
    ///</summary>
    public Dictionary<string, string> CoreProperties
    {
      get
      {
        if( _package.PartExists( new Uri( "/docProps/core.xml", UriKind.Relative ) ) )
        {
          PackagePart docProps_Core = _package.GetPart( new Uri( "/docProps/core.xml", UriKind.Relative ) );
          XDocument corePropDoc;
          using( TextReader tr = new StreamReader( docProps_Core.GetStream( FileMode.Open, FileAccess.Read ) ) )
            corePropDoc = XDocument.Load( tr, LoadOptions.PreserveWhitespace );

          // Get all of the core properties in this document
          return ( from docProperty in corePropDoc.Root.Elements()
                   select
                     new KeyValuePair<string, string>(
                     string.Format(
                       "{0}:{1}",
                       corePropDoc.Root.GetPrefixOfNamespace( docProperty.Name.Namespace ),
                       docProperty.Name.LocalName ),
                     docProperty.Value ) ).ToDictionary( p => p.Key, v => v.Value );
        }

        return new Dictionary<string, string>();
      }
    }

    /// <summary>
    /// Get the Text of this document.
    /// </summary>
    /// <example>
    /// Write to Console the Text from this document.
    /// <code>
    /// // Load a document
    /// DocX document = DocX.Load(@"C:\Example\Test.docx");
    ///
    /// // Get the text of this document.
    /// string text = document.Text;
    ///
    /// // Write the text of this document to Console.
    /// Console.Write(text);
    ///
    /// // Wait for the user to press a key before closing the console window.
    /// Console.ReadKey();
    /// </code>
    /// </example>
    public string Text
    {
      get
      {
        return HelperFunctions.GetText( Xml );
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

    public override List<List> Lists
    {
      get
      {
        var l = base.Lists;
        l.ForEach( x => x.Items.ForEach( i => i.PackagePart = this.PackagePart ) );
        return l;
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

    /// <summary>
    /// Get the Footnotes of this document
    /// </summary>
    public IEnumerable<string> FootnotesText
    {
      get
      {
        foreach( var note in _footnotes.Root.Elements( w + "footnote" ) )
          yield return HelperFunctions.GetText( note );
      }
    }

    /// <summary>
    /// Get the Endnotes of this document
    /// </summary>
    public IEnumerable<string> EndnotesText
    {
      get
      {
        foreach( var note in _endnotes.Root.Elements( w + "endnote" ) )
          yield return HelperFunctions.GetText( note );
      }
    }

    public BookmarkCollection Bookmarks
    {
      get
      {
        var bookmarks = new BookmarkCollection();
        for( int i = 0; i < this.Paragraphs.Count; ++i )
        {
          bookmarks.AddRange( this.Paragraphs[ i ].GetBookmarks() );
        }
        return bookmarks;
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Returns the type of editing protection imposed on this document.
    /// </summary>
    /// <returns>The type of editing protection imposed on this document.</returns>
    /// <example>
    /// <code>
    /// Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Make sure the document is protected before checking the protection type.
    ///     if (document.isProtected)
    ///     {
    ///         EditRestrictions protection = document.GetProtectionType();
    ///         Console.WriteLine("Document is protected using " + protection.ToString());
    ///     }
    ///
    ///     else
    ///         Console.WriteLine("Document is not protected.");
    ///
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AddProtection"/>
    /// <seealso cref="RemoveProtection"/>
    /// <seealso cref="isProtected"/>
    public EditRestrictions GetProtectionType()
    {
      if( isProtected )
      {
        XElement documentProtection = _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).FirstOrDefault();
        string edit_type = documentProtection.Attribute( XName.Get( "edit", w.NamespaceName ) ).Value;
        return ( EditRestrictions )Enum.Parse( typeof( EditRestrictions ), edit_type );
      }

      return EditRestrictions.none;
    }

    /// <summary>
    /// Add editing protection to this document. 
    /// </summary>
    /// <param name="er">The type of protection to add to this document.</param>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Allow no editing, only the adding of comment.
    ///     document.AddProtection(EditRestrictions.comments);
    ///     
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="RemoveProtection"/>
    /// <seealso cref="GetProtectionType"/>
    /// <seealso cref="isProtected"/>
    public void AddProtection( EditRestrictions er )
    {
      // Call remove protection before adding a new protection element.
      RemoveProtection();

      if( er == EditRestrictions.none )
        return;

      var documentProtection = new XElement( XName.Get( "documentProtection", w.NamespaceName ) );
      documentProtection.Add( new XAttribute( XName.Get( "edit", w.NamespaceName ), er.ToString() ) );
      documentProtection.Add( new XAttribute( XName.Get( "enforcement", w.NamespaceName ), "1" ) );

      _settings.Root.AddFirst( documentProtection );
    }

    /// <summary>
    /// Remove editing protection from this document.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a new document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Remove any editing restrictions that are imposed on this document.
    ///     document.RemoveProtection();
    ///
    ///     // Save the document.
    ///     document.Save();
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AddProtection(EditRestrictions)"/>
    /// <seealso cref="GetProtectionType"/>
    /// <seealso cref="isProtected"/>
    public void RemoveProtection()
    {
      // Remove every node of type documentProtection.
      _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).Remove();
    }

    /// <summary>
    /// Insert the contents of another document at the end of this document. 
    /// </summary>
    /// <param name="remote_document">The document to insert at the end of this document.</param>
    /// <param name="append">When true, document is added at the end. If False, document is added at the beginning.</param>
    /// <example>
    /// Create a new document and insert an old document into it.
    /// <code>
    /// // Create a new document.
    /// using (DocX newDocument = DocX.Create(@"NewDocument.docx"))
    /// {
    ///     // Load an old document.
    ///     using (DocX oldDocument = DocX.Load(@"OldDocument.docx"))
    ///     {
    ///         // Insert the old document into the new document.
    ///         newDocument.InsertDocument(oldDocument);
    ///
    ///         // Save the new document.
    ///         newDocument.Save();
    ///     }// Release the old document from memory.
    /// }// Release the new document from memory.
    /// </code>
    /// <remarks>
    /// If the document being inserted contains Images, CustomProperties and or custom styles, these will be correctly inserted into the new document. In the case of Images, new ID's are generated for the Images being inserted to avoid ID conflicts. CustomProperties with the same name will be ignored not replaced.
    /// </remarks>
    /// </example>
    public void InsertDocument( DocX remote_document, bool append = true )
    {
      // We don't want to effect the origional XDocument, so create a new one from the old one.
      var remote_mainDoc = new XDocument( remote_document._mainDoc );

      XDocument remote_footnotes = null;
      if( remote_document._footnotes != null )
      {
        remote_footnotes = new XDocument( remote_document._footnotes );
      }

      XDocument remote_endnotes = null;
      if( remote_document._endnotes != null )
      {
        remote_endnotes = new XDocument( remote_document._endnotes );
      }

      // Remove all header and footer references.
      remote_mainDoc.Descendants( XName.Get( "headerReference", w.NamespaceName ) ).Remove();
      remote_mainDoc.Descendants( XName.Get( "footerReference", w.NamespaceName ) ).Remove();

      // Get the body of the remote document.
      var remote_body = remote_mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );

      // Every file that is missing from the local document will have to be copied, every file that already exists will have to be merged.
      var ppc = remote_document._package.GetParts();

      var ignoreContentTypes = new List<string>
            {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
                "application/vnd.openxmlformats-package.core-properties+xml",
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                ContentTypeApplicationRelationShipXml
            };

      var imageContentTypes = new List<string>
            {
                "image/jpeg",
                "image/jpg",
                "image/png",
                "image/bmp",
                "image/gif",
                "image/tiff",
                "image/icon",
                "image/pcx",
                "image/emf",
                "image/wmf"
            };
      // Check if each PackagePart pp exists in this document.
      foreach( PackagePart remote_pp in ppc )
      {
        if( ignoreContentTypes.Contains( remote_pp.ContentType ) || imageContentTypes.Contains( remote_pp.ContentType ) )
          continue;

        // If this external PackagePart already exits then we must merge them.
        if( _package.PartExists( remote_pp.Uri ) )
        {
          var local_pp = _package.GetPart( remote_pp.Uri );
          switch( remote_pp.ContentType )
          {
            case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
              merge_customs( remote_pp, local_pp, remote_mainDoc );
              break;

            // Merge footnotes/endnotes before merging styles, then set the remote_footnotes to the just updated footnotes
            case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
              remote_footnotes = _footnotes;
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
              remote_endnotes = _endnotes;
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
              merge_styles( remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes, remote_endnotes );
              break;

            // Merges Styles after merging the footnotes, so the changes will be applied to the correct document/footnotes.
            case "application/vnd.ms-word.stylesWithEffects+xml":
              merge_styles( remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes, remote_endnotes );
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
              merge_fonts( remote_pp, local_pp, remote_mainDoc, remote_document );
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
              merge_numbering( remote_pp, local_pp, remote_mainDoc, remote_document );
              break;
          }
        }
        // If this external PackagePart does not exits in the internal document then we can simply copy it.
        else
        {
          var packagePart = clonePackagePart( remote_pp );
          switch( remote_pp.ContentType )
          {
            case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
              _endnotesPart = packagePart;
              _endnotes = remote_endnotes;
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
              _footnotesPart = packagePart;
              _footnotes = remote_footnotes;
              break;

            case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
              _stylesPart = packagePart;
              using( TextReader tr = new StreamReader( _stylesPart.GetStream() ) )
                _styles = XDocument.Load( tr );
              break;

            case "application/vnd.ms-word.stylesWithEffects+xml":
              _stylesWithEffectsPart = packagePart;
              using( TextReader tr = new StreamReader( _stylesWithEffectsPart.GetStream() ) )
                _stylesWithEffects = XDocument.Load( tr );
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
              _fontTablePart = packagePart;
              using( TextReader tr = new StreamReader( _fontTablePart.GetStream() ) )
                _fontTable = XDocument.Load( tr );
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
              _numberingPart = packagePart;
              using( TextReader tr = new StreamReader( _numberingPart.GetStream() ) )
                _numbering = XDocument.Load( tr );
              break;

          }

          clonePackageRelationship( remote_document, remote_pp, remote_mainDoc );
        }
      }

      foreach( var hyperlink_rel in remote_document.PackagePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ) )
      {
        var old_rel_Id = hyperlink_rel.Id;
        var new_rel_Id = this.PackagePart.CreateRelationship( hyperlink_rel.TargetUri, hyperlink_rel.TargetMode, hyperlink_rel.RelationshipType ).Id;
        var hyperlink_refs = remote_mainDoc.Descendants( XName.Get( "hyperlink", w.NamespaceName ) );
        foreach( var hyperlink_ref in hyperlink_refs )
        {
          var a0 = hyperlink_ref.Attribute( XName.Get( "id", r.NamespaceName ) );
          if( a0 != null && a0.Value == old_rel_Id )
          {
            a0.SetValue( new_rel_Id );
          }
        }
      }

      foreach( var oleObject_rel in remote_document.PackagePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" ) )
      {
        var oldRelationshipID = oleObject_rel.Id;
        var newRelationshipID = this.PackagePart.CreateRelationship( oleObject_rel.TargetUri, oleObject_rel.TargetMode, oleObject_rel.RelationshipType ).Id;
        var references = remote_mainDoc.Descendants( XName.Get( "OLEObject", "urn:schemas-microsoft-com:office:office" ) );

        foreach( var reference in references )
        {
          var attribute = reference.Attribute( XName.Get( "id", r.NamespaceName ) );
          if( attribute != null && attribute.Value == oldRelationshipID )
            attribute.SetValue( newRelationshipID );
        }
      }

      foreach( PackagePart remote_pp in ppc )
      {
        if( imageContentTypes.Contains( remote_pp.ContentType ) )
        {
          merge_images( remote_pp, remote_document, remote_mainDoc, remote_pp.ContentType );
        }
      }

      int id = 0;
      var local_docPrs = _mainDoc.Root.Descendants( XName.Get( "docPr", wp.NamespaceName ) );
      foreach( var local_docPr in local_docPrs )
      {
        var a_id = local_docPr.Attribute( XName.Get( "id" ) );
        int a_id_value;
        if( a_id != null && int.TryParse( a_id.Value, out a_id_value ) )
        {
          if( a_id_value > id )
          {
            id = a_id_value;
          }
        }
      }
      id++;

      // docPr must be sequential
      var docPrs = remote_body.Descendants( XName.Get( "docPr", wp.NamespaceName ) );
      foreach( var docPr in docPrs )
      {
        docPr.SetAttributeValue( XName.Get( "id" ), id );
        id++;
      }

      // Add the remote documents contents to this document.
      var local_body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
      if( append )
      {
        local_body.Add( remote_body.Elements() );
      }
      else
      {
        local_body.AddFirst( remote_body.Elements() );
      }

      // Copy any missing root attributes to the local document.
      foreach( XAttribute a in remote_mainDoc.Root.Attributes() )
      {
        if( _mainDoc.Root.Attribute( a.Name ) == null )
        {
          _mainDoc.Root.SetAttributeValue( a.Name, a.Value );
        }
      }
    }

    /// <summary>
    /// Insert a new Table at the end of this document.
    /// </summary>
    /// <param name="columnCount">The number of columns to create.</param>
    /// <param name="rowCount">The number of rows to create.</param>
    /// <returns>A new Table.</returns>
    /// <example>
    /// Insert a new Table with 2 columns and 3 rows, at the end of a document.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"C:\Example\Test.docx"))
    /// {
    ///     // Create a new Table with 2 columns and 3 rows.
    ///     Table newTable = document.InsertTable(2, 3);
    ///
    ///     // Set the design of this Table.
    ///     newTable.Design = TableDesign.LightShadingAccent2;
    ///
    ///     // Set the column names.
    ///     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
    ///     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
    ///
    ///     // Fill row 1
    ///     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
    ///     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
    ///
    ///     // Fill row 2
    ///     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
    ///     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
    ///
    ///     // Save all changes made to document b.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public new Table InsertTable( int rowCount, int columnCount )
    {
      if( rowCount < 1 || columnCount < 1 )
        throw new ArgumentOutOfRangeException( "Row and Column count must be greater than zero." );

      var t = base.InsertTable( rowCount, columnCount );
      t.PackagePart = this.PackagePart;
      return t;
    }

    public Table AddTable( int rowCount, int columnCount )
    {
      if( rowCount < 1 || columnCount < 1 )
        throw new ArgumentOutOfRangeException( "Row and Column count must be greater than zero." );

      var t = new Table( this, HelperFunctions.CreateTable( rowCount, columnCount ) );
      t.PackagePart = this.PackagePart;
      return t;
    }

    /// <summary>
    /// Insert a Table into this document. The Table's source can be a completely different document.
    /// </summary>
    /// <param name="t">The Table to insert.</param>
    /// <param name="index">The index to insert this Table at.</param>
    /// <returns>The Table now associated with this document.</returns>
    /// <example>
    /// Extract a Table from document a and insert it into document b, at index 10.
    /// <code>
    /// // Place holder for a Table.
    /// Table t;
    ///
    /// // Load document a.
    /// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
    /// {
    ///     // Get the first Table from this document.
    ///     t = documentA.Tables[0];
    /// }
    ///
    /// // Load document b.
    /// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
    /// {
    ///     /* 
    ///      * Insert the Table that was extracted from document a, into document b. 
    ///      * This creates a new Table that is now associated with document b.
    ///      */
    ///     Table newTable = documentB.InsertTable(10, t);
    ///
    ///     // Save all changes made to document b.
    ///     documentB.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public new Table InsertTable( int index, Table t )
    {
      var t2 = base.InsertTable( index, t );
      t2.PackagePart = this.PackagePart;
      return t2;
    }

    /// <summary>
    /// Insert a Table into this document. The Table's source can be a completely different document.
    /// </summary>
    /// <param name="t">The Table to insert.</param>
    /// <returns>The Table now associated with this document.</returns>
    /// <example>
    /// Extract a Table from document a and insert it at the end of document b.
    /// <code>
    /// // Place holder for a Table.
    /// Table t;
    ///
    /// // Load document a.
    /// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
    /// {
    ///     // Get the first Table from this document.
    ///     t = documentA.Tables[0];
    /// }
    ///
    /// // Load document b.
    /// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
    /// {
    ///     /* 
    ///      * Insert the Table that was extracted from document a, into document b. 
    ///      * This creates a new Table that is now associated with document b.
    ///      */
    ///     Table newTable = documentB.InsertTable(t);
    ///
    ///     // Save all changes made to document b.
    ///     documentB.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public new Table InsertTable( Table t )
    {
      t = base.InsertTable( t );
      t.PackagePart = this.PackagePart;
      return t;
    }

    /// <summary>
    /// Insert a new Table at the end of this document.
    /// </summary>
    /// <param name="columnCount">The number of columns to create.</param>
    /// <param name="rowCount">The number of rows to create.</param>
    /// <param name="index">The index to insert this Table at.</param>
    /// <returns>A new Table.</returns>
    /// <example>
    /// Insert a new Table with 2 columns and 3 rows, at index 37 in this document.
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Create a new Table with 3 rows and 2 columns. Insert this Table at index 37.
    ///     Table newTable = document.InsertTable(37, 3, 2);
    ///
    ///     // Set the design of this Table.
    ///     newTable.Design = TableDesign.LightShadingAccent3;
    ///
    ///     // Set the column names.
    ///     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
    ///     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
    ///
    ///     // Fill row 1
    ///     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
    ///     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
    ///
    ///     // Fill row 2
    ///     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
    ///     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
    ///
    ///     // Save all changes made to document b.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public new Table InsertTable( int index, int rowCount, int columnCount )
    {
      if( rowCount < 1 || columnCount < 1 )
        throw new ArgumentOutOfRangeException( "Row and Column count must be greater than zero." );

      var t = base.InsertTable( index, rowCount, columnCount );
      t.PackagePart = this.PackagePart;
      return t;
    }

    /// <summary>
    /// Creates a document using a Stream.
    /// </summary>
    /// <param name="stream">The Stream to create the document from.</param>
    /// <param name="documentType"></param>
    /// <returns>Returns a DocX object which represents the document.</returns>
    /// <example>
    /// Creating a document from a FileStream.
    /// <code>
    /// // Use a FileStream fs to create a new document.
    /// using(FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Create))
    /// {
    ///     // Load the document using fs
    ///     using (DocX document = DocX.Create(fs))
    ///     {
    ///         // Do something with the document here.
    ///
    ///         // Save all changes made to this document.
    ///         document.Save();
    ///     }// Release this document from memory.
    /// }
    /// </code>
    /// </example>
    /// <example>
    /// Creating a document in a SharePoint site.
    /// <code>
    /// using(SPSite mySite = new SPSite("http://server/sites/site"))
    /// {
    ///     // Open a connection to the SharePoint site
    ///     using(SPWeb myWeb = mySite.OpenWeb())
    ///     {
    ///         // Create a MemoryStream ms.
    ///         using (MemoryStream ms = new MemoryStream())
    ///         {
    ///             // Create a document using ms.
    ///             using (DocX document = DocX.Create(ms))
    ///             {
    ///                 // Do something with the document here.
    ///
    ///                 // Save all changes made to this document.
    ///                 document.Save();
    ///             }// Release this document from memory
    ///
    ///             // Add the document to the SharePoint site
    ///             web.Files.Add("filename", ms.ToArray(), true);
    ///         }
    ///     }
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Load(string)"/>
    /// <seealso cref="DocX.Save()"/>
    public static DocX Create( Stream stream, DocumentTypes documentType = DocumentTypes.Document )
    {
      var document = CreateDocument( documentType );
      document._stream = stream;
      return document;
    }

    /// <summary>
    /// Creates a document using a fully qualified or relative filename.
    /// </summary>
    /// <param name="filename">The fully qualified or relative filename.</param>
    /// <param name="documentType"></param>
    /// <returns>Returns a DocX object which represents the document.</returns>
    /// <example>
    /// <code>
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Create(@"..\Test.docx"))
    /// {
    ///     // Do something with the document here.
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory
    /// </code>
    /// <code>
    /// // Create a document using a relative filename.
    /// using (DocX document = DocX.Create(@"..\Test.docx"))
    /// {
    ///     // Do something with the document here.
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory
    /// </code>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Load(string)"/>
    /// <seealso cref="DocX.Save()"/>
    /// </example>
    public static DocX Create( string filename, DocumentTypes documentType = DocumentTypes.Document )
    {
      var document = CreateDocument( documentType );
      document._filename = filename;
      return document;
    }

    /// <summary>
    /// Loads a document into a DocX object using a Stream.
    /// </summary>
    /// <param name="stream">The Stream to load the document from.</param>
    /// <returns>
    /// Returns a DocX object which represents the document.
    /// </returns>
    /// <example>
    /// Loading a document from a FileStream.
    /// <code>
    /// // Open a FileStream fs to a document.
    /// using (FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Open))
    /// {
    ///     // Load the document using fs.
    ///     using (DocX document = DocX.Load(fs))
    ///     {
    ///         // Do something with the document here.
    ///            
    ///         // Save all changes made to the document.
    ///         document.Save();
    ///     }// Release this document from memory.
    /// }
    /// </code>
    /// </example>
    /// <example>
    /// Loading a document from a SharePoint site.
    /// <code>
    /// // Get the SharePoint site that you want to access.
    /// using (SPSite mySite = new SPSite("http://server/sites/site"))
    /// {
    ///     // Open a connection to the SharePoint site
    ///     using (SPWeb myWeb = mySite.OpenWeb())
    ///     {
    ///         // Grab a document stored on this site.
    ///         SPFile file = web.GetFile("Source_Folder_Name/Source_File");
    ///
    ///         // DocX.Load requires a Stream, so open a Stream to this document.
    ///         Stream str = new MemoryStream(file.OpenBinary());
    ///
    ///         // Load the file using the Stream str.
    ///         using (DocX document = DocX.Load(str))
    ///         {
    ///             // Do something with the document here.
    ///
    ///             // Save all changes made to the document.
    ///             document.Save();
    ///         }// Release this document from memory.
    ///     }
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="DocX.Load(string)"/>
    /// <seealso cref="DocX.Save()"/>
    public static DocX Load( Stream stream )
    {
      var ms = new MemoryStream();

      stream.Position = 0;
      var data = new byte[ stream.Length ];
      stream.Read( data, 0, ( int )stream.Length );
      ms.Write( data, 0, ( int )stream.Length );

      // Open the docx package
      var package = Package.Open( ms, FileMode.Open, FileAccess.ReadWrite );

      var document = DocX.PostLoad( ref package );
      document._package = package;
      document._memoryStream = ms;
      document._stream = stream;
      return document;
    }

    /// <summary>
    /// Loads a document into a DocX object using a fully qualified or relative filename.
    /// </summary>
    /// <param name="filename">The fully qualified or relative filename.</param>
    /// <returns>
    /// Returns a DocX object which represents the document.
    /// </returns>
    /// <example>
    /// <code>
    /// // Load a document using its fully qualified filename
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Do something with the document here
    ///
    ///     // Save all changes made to document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// <code>
    /// // Load a document using its relative filename.
    /// using(DocX document = DocX.Load(@"..\..\Test.docx"))
    /// { 
    ///     // Do something with the document here.
    ///                
    ///     // Save all changes made to document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Save()"/>
    /// </example>
    public static DocX Load( string filename )
    {
      if( !File.Exists( filename ) )
        throw new FileNotFoundException( string.Format( "File could not be found {0}", filename ) );

      var ms = new MemoryStream();

      using( FileStream fs = new FileStream( filename, FileMode.Open, FileAccess.Read, FileShare.Read ) )
      {
        CopyStream( fs, ms );
      }

      // Open the docx package
      var package = Package.Open( ms, FileMode.Open, FileAccess.ReadWrite );

      var document = PostLoad( ref package );
      document._package = package;
      document._filename = filename;
      document._memoryStream = ms;

      return document;
    }

    ///<summary>
    /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
    ///</summary>
    ///<param name="templateFilePath">The path to the document template file.</param>
    ///<exception cref="FileNotFoundException">The document template file not found.</exception>
    public void ApplyTemplate( string templateFilePath )
    {
      ApplyTemplate( templateFilePath, true );
    }

    ///<summary>
    /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
    ///</summary>
    ///<param name="templateFilePath">The path to the document template file.</param>
    ///<param name="includeContent">Whether to copy the document template text content to document.</param>
    ///<exception cref="FileNotFoundException">The document template file not found.</exception>
    public void ApplyTemplate( string templateFilePath, bool includeContent )
    {
      if( !File.Exists( templateFilePath ) )
      {
        throw new FileNotFoundException( string.Format( "File could not be found {0}", templateFilePath ) );
      }
      using( FileStream packageStream = new FileStream( templateFilePath, FileMode.Open, FileAccess.Read ) )
      {
        ApplyTemplate( packageStream, includeContent );
      }
    }

    ///<summary>
    /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
    ///</summary>
    ///<param name="templateStream">The stream of the document template file.</param>
    public void ApplyTemplate( Stream templateStream )
    {
      ApplyTemplate( templateStream, true );
    }

    ///<summary>
    /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
    ///</summary>
    ///<param name="templateStream">The stream of the document template file.</param>
    ///<param name="includeContent">Whether to copy the document template text content to document.</param>
    public void ApplyTemplate( Stream templateStream, bool includeContent )
    {
      var templatePackage = Package.Open( templateStream );
      try
      {
        PackagePart documentPart = null;
        XDocument documentDoc = null;
        foreach( PackagePart packagePart in templatePackage.GetParts() )
        {
          switch( packagePart.Uri.ToString() )
          {
            case "/word/document.xml":
              documentPart = packagePart;
              using( XmlReader xr = XmlReader.Create( packagePart.GetStream( FileMode.Open, FileAccess.Read ) ) )
              {
                documentDoc = XDocument.Load( xr );
              }
              break;
            case "/_rels/.rels":
              if( !_package.PartExists( packagePart.Uri ) )
              {
                _package.CreatePart( packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption );
              }
              var globalRelsPart = _package.GetPart( packagePart.Uri );
              using(
                var tr = new StreamReader(
                  packagePart.GetStream( FileMode.Open, FileAccess.Read ), Encoding.UTF8 ) )
              {
                using(
                  var tw = new StreamWriter(
                    new PackagePartStream( globalRelsPart.GetStream( FileMode.Create, FileAccess.Write ) ), Encoding.UTF8 ) )
                {
                  tw.Write( tr.ReadToEnd() );
                }
              }
              break;
            case "/word/_rels/document.xml.rels":
              break;
            default:
              if( !_package.PartExists( packagePart.Uri ) )
              {
                _package.CreatePart( packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption );
              }
              var packagePartEncoding = Encoding.Default;
              if( packagePart.Uri.ToString().EndsWith( ".xml" ) || packagePart.Uri.ToString().EndsWith( ".rels" ) )
              {
                packagePartEncoding = Encoding.UTF8;
              }
              var nativePart = _package.GetPart( packagePart.Uri );
              using(
                var tr = new StreamReader(
                  packagePart.GetStream( FileMode.Open, FileAccess.Read ), packagePartEncoding ) )
              {
                using(
                  var tw = new StreamWriter(
                    new PackagePartStream( nativePart.GetStream( FileMode.Create, FileAccess.Write ) ), tr.CurrentEncoding ) )
                {
                  tw.Write( tr.ReadToEnd() );
                }
              }
              break;
          }
        }
        if( documentPart != null )
        {
          string mainContentType = documentPart.ContentType.Replace( "template.main", "document.main" );
          if( _package.PartExists( documentPart.Uri ) )
          {
            _package.DeletePart( documentPart.Uri );
          }
          var documentNewPart = _package.CreatePart(
            documentPart.Uri, mainContentType, documentPart.CompressionOption );
          using( XmlWriter xw = XmlWriter.Create( new PackagePartStream( documentNewPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            documentDoc.WriteTo( xw );
          }
          foreach( PackageRelationship documentPartRel in documentPart.GetRelationships() )
          {
            documentNewPart.CreateRelationship(
              documentPartRel.TargetUri,
              documentPartRel.TargetMode,
              documentPartRel.RelationshipType,
              documentPartRel.Id );
          }
          this.PackagePart = documentNewPart;
          _mainDoc = documentDoc;
          PopulateDocument( this, templatePackage );

          // DragonFire: I added next line and recovered ApplyTemplate method. 
          // I do it, becouse  PopulateDocument(...) writes into field "settingsPart" the part of Template's package 
          //  and after line "templatePackage.Close();" in finally, field "settingsPart" becomes not available and method "Save" throw an exception...
          // That's why I recreated settingsParts and unlinked it from Template's package =)
          _settingsPart = HelperFunctions.CreateOrGetSettingsPart( _package );
        }
        if( !includeContent )
        {
          foreach( Paragraph paragraph in this.Paragraphs )
          {
            paragraph.Remove( false );
          }
        }
      }
      finally
      {
        _package.Flush();
        var documentRelsPart = _package.GetPart( new Uri( "/word/_rels/document.xml.rels", UriKind.Relative ) );
        using( TextReader tr = new StreamReader( documentRelsPart.GetStream( FileMode.Open, FileAccess.Read ) ) )
        {
          tr.Read();
        }
        templatePackage.Close();
        PopulateDocument( Document, _package );
      }
    }

    /// <summary>
    /// Add an Image into this document from a fully qualified or relative filename.
    /// </summary>
    /// <param name="filename">The fully qualified or relative filename.</param>
    /// <param name="contentType">MIME type of image, guessed if not given.</param>
    /// <returns>An Image file.</returns>
    /// <example>
    /// Add an Image into this document from a fully qualified filename.
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Add an Image from a file.
    ///     document.AddImage(@"C:\Example\Image.png");
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <seealso cref="AddImage(Stream, string)"/>
    /// <seealso cref="Paragraph.InsertPicture"/>
    public Image AddImage( string filename )
    {
      string contentType = "";

      // The extension this file has will be taken to be its format.
      switch( Path.GetExtension( filename ) )
      {
        case ".tiff":
          contentType = "image/tif";
          break;
        case ".tif":
          contentType = "image/tif";
          break;
        case ".png":
          contentType = "image/png";
          break;
        case ".bmp":
          contentType = "image/png";
          break;
        case ".gif":
          contentType = "image/gif";
          break;
        case ".jpg":
          contentType = "image/jpg";
          break;
        case ".jpeg":
          contentType = "image/jpeg";
          break;
        default:
          contentType = "image/jpg";
          break;
      }

      return AddImage( filename as object, contentType );
    }

    /// <summary>
    /// Add an Image into this document from a Stream.
    /// </summary>
    /// <param name="stream">A Stream stream.</param>
    /// <param name="contentType">MIME type of image.</param>
    /// <returns>An Image file.</returns>
    /// <example>
    /// Add an Image into a document using a Stream. 
    /// <code>
    /// // Open a FileStream fs to an Image.
    /// using (FileStream fs = new FileStream(@"C:\Example\Image.jpg", FileMode.Open))
    /// {
    ///     // Load a document.
    ///     using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    ///     {
    ///         // Add an Image from a filestream fs.
    ///         document.AddImage(fs);
    ///
    ///         // Save all changes made to this document.
    ///         document.Save();
    ///     }// Release this document from memory.
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="AddImage(string)"/>
    /// <seealso cref="Paragraph.InsertPicture"/>
    public Image AddImage( Stream stream, string contentType = "image/jpeg" )
    {
      return AddImage( stream as object, contentType );
    }

    /// <summary>
    /// Adds a hyperlink with a uri to a document and creates a Paragraph which uses it.
    /// </summary>
    /// <param name="text">The text as displayed by the hyperlink.</param>
    /// <param name="uri">The hyperlink itself.</param>
    /// <returns>Returns a hyperlink with a uri that can be inserted into a Paragraph.</returns>
    /// <example>
    /// Adds a hyperlink to a document and creates a Paragraph which uses it.
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
    public Hyperlink AddHyperlink( string text, Uri uri )
    {
      return this.AddHyperlink( text, uri, null );     
    }

    /// <summary>
    /// Adds a hyperlink with an anchor to a document and creates a Paragraph which uses it.
    /// </summary>
    /// <param name="text">The text as displayed by the hyperlink.</param>
    /// <param name="anchor">The anchor to a bookmark.</param>
    /// <returns>Returns a hyperlink with an anchor that can be inserted into a Paragraph.</returns>
    public Hyperlink AddHyperlink( string text, string anchor )
    {
      return this.AddHyperlink( text, null, anchor );
    }

    /// <summary>
    /// Adds three new Headers to this document. One for the first page, one for odd pages and one for even pages.
    /// </summary>
    /// <example>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add header support to this document.
    ///     document.AddHeaders();
    ///
    ///     // Get a collection of all headers in this document.
    ///     Headers headers = document.Headers;
    ///
    ///     // The header used for the first page of this document.
    ///     Header first = headers.first;
    ///
    ///     // The header used for odd pages of this document.
    ///     Header odd = headers.odd;
    ///
    ///     // The header used for even pages of this document.
    ///     Header even = headers.even;
    ///
    ///     // Force the document to use a different header for first, odd and even pages.
    ///     document.DifferentFirstPage = true;
    ///     document.DifferentOddAndEvenPages = true;
    ///
    ///     // Content can be added to the Headers in the same manor that it would be added to the main document.
    ///     Paragraph p = first.InsertParagraph();
    ///     p.Append("This is the first pages header.");
    ///
    ///     // Save all changes to this document.
    ///     document.Save();    
    /// }// Release this document from memory.
    /// </example>
    public void AddHeaders()
    {
      this.AddHeadersOrFooters( true );

      _headers.Odd = Document.GetHeaderByType( "default" );
      _headers.Even = Document.GetHeaderByType( "even" );
      _headers.First = Document.GetHeaderByType( "first" );
    }

    /// <summary>
    /// Adds three new Footers to this document. One for the first page, one for odd pages and one for even pages.
    /// </summary>
    /// <example>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add footer support to this document.
    ///     document.AddFooters();
    ///
    ///     // Get a collection of all footers in this document.
    ///     Footers footers = document.Footers;
    ///
    ///     // The footer used for the first page of this document.
    ///     Footer first = footers.first;
    ///
    ///     // The footer used for odd pages of this document.
    ///     Footer odd = footers.odd;
    ///
    ///     // The footer used for even pages of this document.
    ///     Footer even = footers.even;
    ///
    ///     // Force the document to use a different footer for first, odd and even pages.
    ///     document.DifferentFirstPage = true;
    ///     document.DifferentOddAndEvenPages = true;
    ///
    ///     // Content can be added to the Footers in the same manor that it would be added to the main document.
    ///     Paragraph p = first.InsertParagraph();
    ///     p.Append("This is the first pages footer.");
    ///
    ///     // Save all changes to this document.
    ///     document.Save();    
    /// }// Release this document from memory.
    /// </example>
    public void AddFooters()
    {
      AddHeadersOrFooters( false );

      _footers.Odd = Document.GetFooterByType( "default" );
      _footers.Even = Document.GetFooterByType( "even" );
      _footers.First = Document.GetFooterByType( "first" );
    }

    /// <summary>
    /// Save this document back to the location it was loaded from.
    /// </summary>
    /// <example>
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Add an Image from a file.
    ///     document.AddImage(@"C:\Example\Image.jpg");
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <seealso cref="DocX.SaveAs(string)"/>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Load(string)"/> 
    /// <!-- 
    /// Bug found and fixed by krugs525 on August 12 2009.
    /// Use TFS compare to see exact code change.
    /// -->
    public void Save()
    {
      var headers = Headers;

      // Save the main document
      using( TextWriter tw = new StreamWriter( new PackagePartStream( this.PackagePart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        _mainDoc.Save( tw, SaveOptions.None );
      }

      if( (_settings == null) || !this.isProtected )
      {
        using( TextReader textReader = new StreamReader( _settingsPart.GetStream() ) )
        {
          _settings = XDocument.Load( textReader );
        }
      }

      var body = _mainDoc.Root.Element( w + "body" );
      var sectPr = body.Descendants( w + "sectPr" ).FirstOrDefault();

      if( sectPr != null )
      {
        var evenHeaderRef =
        (
            from e in _mainDoc.Descendants( w + "headerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "even", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( evenHeaderRef != null )
        {
          var even = headers.Even.Xml;

          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( evenHeaderRef ).TargetUri
          );

          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                even
            ).Save( tw, SaveOptions.None );
          }
        }

        var oddHeaderRef =
        (
            from e in _mainDoc.Descendants( w + "headerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "default", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( oddHeaderRef != null )
        {
          var odd = headers.Odd.Xml;

          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( oddHeaderRef ).TargetUri
          );

          // Save header1
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                odd
            ).Save( tw, SaveOptions.None );
          }
        }

        var firstHeaderRef =
        (
            from e in _mainDoc.Descendants( w + "headerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "first", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( firstHeaderRef != null )
        {
          var first = headers.First.Xml;
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( firstHeaderRef ).TargetUri
          );

          // Save header3
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                first
            ).Save( tw, SaveOptions.None );
          }
        }

        var oddFooterRef =
        (
            from e in _mainDoc.Descendants( w + "footerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "default", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( oddFooterRef != null )
        {
          var odd = _footers.Odd.Xml;
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( oddFooterRef ).TargetUri
          );

          // Save header1
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                odd
            ).Save( tw, SaveOptions.None );
          }
        }

        var evenFooterRef =
        (
            from e in _mainDoc.Descendants( w + "footerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "even", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( evenFooterRef != null )
        {
          var even = _footers.Even.Xml;
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( evenFooterRef ).TargetUri
          );

          // Save header2
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                even
            ).Save( tw, SaveOptions.None );
          }
        }

        var firstFooterRef =
        (
             from e in _mainDoc.Descendants( w + "footerReference" )
             let type = e.Attribute( w + "type" )
             where type != null && type.Value.Equals( "first", StringComparison.CurrentCultureIgnoreCase )
             select e.Attribute( r + "id" ).Value
        ).LastOrDefault();

        if( firstFooterRef != null )
        {
          var first = _footers.First.Xml;
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( firstFooterRef ).TargetUri
          );

          // Save header3
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                first
            ).Save( tw, SaveOptions.None );
          }
        }

        // Save the settings document.
        using( TextWriter tw = new StreamWriter( new PackagePartStream( _settingsPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
        {
          _settings.Save( tw, SaveOptions.None );
        }

        if( _endnotesPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _endnotesPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _endnotes.Save( tw, SaveOptions.None );
          }
        }

        if( _footnotesPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _footnotesPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _footnotes.Save( tw, SaveOptions.None );
          }
        }

        if( _stylesPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _stylesPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _styles.Save( tw, SaveOptions.None );
          }
        }

        if( _stylesWithEffectsPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _stylesWithEffectsPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _stylesWithEffects.Save( tw, SaveOptions.None );
          }
        }

        if( _numberingPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _numberingPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _numbering.Save( tw, SaveOptions.None );
          }
        }

        if( _fontTablePart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _fontTablePart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _fontTable.Save( tw, SaveOptions.None );
          }
        }
      }

      // Close the document so that it can be saved.
      _package.Flush();

      #region Save this document back to a file or stream, that was specified by the user at save time.
      if( _filename != null )
      {
        using( FileStream fs = new FileStream( _filename, FileMode.Create ) )
        {
          if( _memoryStream.CanSeek )
          {
            // Write to the beginning of the stream
            _memoryStream.Position = 0;
            CopyStream( _memoryStream, fs );
          }
          else
            fs.Write( _memoryStream.ToArray(), 0, ( int )_memoryStream.Length );
        }
      }
      else if( _stream.CanSeek )  //Check if stream can be seeked to support System.Web.HttpResponseStream.
      {
        // Set the length of this stream to 0
        _stream.SetLength( 0 );

        // Write to the beginning of the stream
        _stream.Position = 0;

        _memoryStream.WriteTo( _stream );
        _memoryStream.Flush();
      }
      #endregion
    }

    /// <summary>
    /// Save this document to a file.
    /// </summary>
    /// <param name="filename">The filename to save this document as.</param>
    /// <example>
    /// Load a document from one file and save it to another.
    /// <code>
    /// // Load a document using its fully qualified filename.
    /// DocX document = DocX.Load(@"C:\Example\Test1.docx");
    ///
    /// // Insert a new Paragraph
    /// document.InsertParagraph("Hello world!", false);
    ///
    /// // Save the document to a new location.
    /// document.SaveAs(@"C:\Example\Test2.docx");
    /// </code>
    /// </example>
    /// <example>
    /// Load a document from a Stream and save it to a file.
    /// <code>
    /// DocX document;
    /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
    /// {
    ///     // Load a document using a stream.
    ///     document = DocX.Load(fs1);
    ///
    ///     // Insert a new Paragraph
    ///     document.InsertParagraph("Hello world again!", false);
    /// }
    ///    
    /// // Save the document to a new location.
    /// document.SaveAs(@"C:\Example\Test2.docx");
    /// </code>
    /// </example>
    /// <seealso cref="DocX.Save()"/>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Load(string)"/>
    public void SaveAs( string filename )
    {
      _filename = filename;
      _stream = null;
      Save();
    }

    /// <summary>
    /// Save this document to a Stream.
    /// </summary>
    /// <param name="stream">The Stream to save this document to.</param>
    /// <example>
    /// Load a document from a file and save it to a Stream.
    /// <code>
    /// // Place holder for a document.
    /// DocX document;
    ///
    /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
    /// {
    ///     // Load a document using a stream.
    ///     document = DocX.Load(fs1);
    ///
    ///     // Insert a new Paragraph
    ///     document.InsertParagraph("Hello world again!", false);
    /// }
    ///
    /// using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
    /// {
    ///     // Save the document to a different stream.
    ///     document.SaveAs(fs2);
    /// }
    ///
    /// // Release this document from memory.
    /// document.Dispose();
    /// </code>
    /// </example>
    /// <example>
    /// Load a document from one Stream and save it to another.
    /// <code>
    /// DocX document;
    /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
    /// {
    ///     // Load a document using a stream.
    ///     document = DocX.Load(fs1);
    ///
    ///     // Insert a new Paragraph
    ///     document.InsertParagraph("Hello world again!", false);
    /// }
    /// 
    /// using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
    /// {
    ///     // Save the document to a different stream.
    ///     document.SaveAs(fs2);
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="DocX.Save()"/>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Load(string)"/>
    public void SaveAs( Stream stream )
    {
      _filename = null;
      _stream = stream;
      Save();
    }

    /// <summary>
    /// Add a core property to this document. If a core property already exists with the same name it will be replaced. Core property names are case insensitive.
    /// </summary>
    ///<param name="propertyName">The property name.</param>
    ///<param name="propertyValue">The property value.</param>
    ///<example>
    /// Add a core properties of each type to a document.
    /// <code>
    /// // Load Example.docx
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // If this document does not contain a core property called 'forename', create one.
    ///     if (!document.CoreProperties.ContainsKey("forename"))
    ///     {
    ///         // Create a new core property called 'forename' and set its value.
    ///         document.AddCoreProperty("forename", "Cathal");
    ///     }
    ///
    ///     // Get this documents core property called 'forename'.
    ///     string forenameValue = document.CoreProperties["forename"];
    ///
    ///     // Print all of the information about this core property to Console.
    ///     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", "forename", forenameValue));
    ///     
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// } // Release this document from memory.
    ///
    /// // Wait for the user to press a key before exiting.
    /// Console.ReadKey();
    /// </code>
    /// </example>
    /// <seealso cref="CoreProperties"/>
    /// <seealso cref="CustomProperty"/>
    /// <seealso cref="CustomProperties"/>
    public void AddCoreProperty( string propertyName, string propertyValue )
    {
      var propertyNamespacePrefix = propertyName.Contains(":") ? propertyName.Split(':')[0] : "cp";
      var propertyLocalName = propertyName.Contains(":") ? propertyName.Split(':')[1] : propertyName;

      // If this document does not contain a coreFilePropertyPart create one.)
      if( !_package.PartExists( new Uri( "/docProps/core.xml", UriKind.Relative ) ) )
        throw new Exception( "Core properties part doesn't exist." );

      XDocument corePropDoc;
      var corePropPart = _package.GetPart( new Uri( "/docProps/core.xml", UriKind.Relative ) );
      using( TextReader tr = new StreamReader( corePropPart.GetStream( FileMode.Open, FileAccess.Read ) ) )
      {
        corePropDoc = XDocument.Load( tr );
      }

      var corePropElement =
        ( from propElement in corePropDoc.Root.Elements()
          where ( propElement.Name.LocalName.Equals( propertyLocalName ) )
          select propElement ).SingleOrDefault();
      if( corePropElement != null )
      {
        corePropElement.SetValue( propertyValue );
      }
      else
      {
        var propertyNamespace = corePropDoc.Root.GetNamespaceOfPrefix( propertyNamespacePrefix );
        corePropDoc.Root.Add( new XElement( XName.Get( propertyLocalName, propertyNamespace.NamespaceName ), propertyValue ) );
      }

      using( TextWriter tw = new StreamWriter( new PackagePartStream( corePropPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        corePropDoc.Save( tw );
      }
      DocX.UpdateCorePropertyValue( this, propertyLocalName, propertyValue );
    }

    /// <summary>
    /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
    /// </summary>
    /// <param name="cp">The CustomProperty to add to this document.</param>
    /// <example>
    /// Add a custom properties of each type to a document.
    /// <code>
    /// // Load Example.docx
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // A CustomProperty called forename which stores a string.
    ///     CustomProperty forename;
    ///
    ///     // If this document does not contain a custom property called 'forename', create one.
    ///     if (!document.CustomProperties.ContainsKey("forename"))
    ///     {
    ///         // Create a new custom property called 'forename' and set its value.
    ///         document.AddCustomProperty(new CustomProperty("forename", "Cathal"));
    ///     }
    ///
    ///     // Get this documents custom property called 'forename'.
    ///     forename = document.CustomProperties["forename"];
    ///
    ///     // Print all of the information about this CustomProperty to Console.
    ///     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", forename.Name, forename.Value));
    ///     
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// } // Release this document from memory.
    ///
    /// // Wait for the user to press a key before exiting.
    /// Console.ReadKey();
    /// </code>
    /// </example>
    /// <seealso cref="CustomProperty"/>
    /// <seealso cref="CustomProperties"/>
    public void AddCustomProperty( CustomProperty cp )
    {
      // If this document does not contain a customFilePropertyPart create one.
      if( !_package.PartExists( new Uri( "/docProps/custom.xml", UriKind.Relative ) ) )
      {
        HelperFunctions.CreateCustomPropertiesPart( this );
      }

      XDocument customPropDoc;
      var customPropPart = _package.GetPart( new Uri( "/docProps/custom.xml", UriKind.Relative ) );
      using( TextReader tr = new StreamReader( customPropPart.GetStream( FileMode.Open, FileAccess.Read ) ) )
      {
        customPropDoc = XDocument.Load( tr, LoadOptions.PreserveWhitespace );
      }

      // Each custom property has a PID, get the highest PID in this document.
      IEnumerable<int> pids =
      (
          from d in customPropDoc.Descendants()
          where d.Name.LocalName == "property"
          select int.Parse( d.Attribute( XName.Get( "pid" ) ).Value )
      );

      int pid = 1;
      if( pids.Count() > 0 )
      {
        pid = pids.Max();
      }

      // Check if a custom property already exists with this name
      var customProperty =
      (
          from d in customPropDoc.Descendants()
          where ( d.Name.LocalName == "property" ) && ( d.Attribute( XName.Get( "name" ) ).Value.Equals( cp.Name, StringComparison.InvariantCultureIgnoreCase ) )
          select d
      ).SingleOrDefault();

      // If a custom property with this name already exists remove it.
      if( customProperty != null )
      {
        customProperty.Remove();
      }

      var propertiesElement = customPropDoc.Element( XName.Get( "Properties", customPropertiesSchema.NamespaceName ) );
      propertiesElement.Add
      (
          new XElement
          (
              XName.Get( "property", customPropertiesSchema.NamespaceName ),
              new XAttribute( "fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" ),
              new XAttribute( "pid", pid + 1 ),
              new XAttribute( "name", cp.Name ),
              new XElement( customVTypesSchema + cp.Type, cp.Value ?? "" )
          )
      );

      // Save the custom properties
      using( TextWriter tw = new StreamWriter( new PackagePartStream( customPropPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        customPropDoc.Save( tw, SaveOptions.None );
      }

      // Refresh all fields in this document which display this custom property.
      DocX.UpdateCustomPropertyValue( this, cp.Name, ( cp.Value ?? "" ).ToString() );
    }

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

    public Paragraph[] InsertParagraphs( string text )
    {
      var textArray = text.Split( '\n' );
      var paragraphs = new List<Paragraph>();
      foreach( var textForParagraph in textArray )
      {
        var p = base.InsertParagraph( text );
        p.PackagePart = this.PackagePart;
        paragraphs.Add( p );
      }
      return paragraphs.ToArray();
    }

    /// <summary>
    /// Create an equation and insert it in the new paragraph
    /// </summary>        
    public override Paragraph InsertEquation( String equation )
    {
      var p = base.InsertEquation( equation );
      p.PackagePart = this.PackagePart;
      return p;
    }

    /// <summary>
    /// Insert a chart in document
    /// </summary>
    public void InsertChart( Chart chart )
    {
      this.InsertChart( chart, null );
    }

    /// <summary>
    /// Insert a chart in document after the specified paragraph
    /// </summary>
    public void InsertChartAfterParagraph( Chart chart, Paragraph paragraph )
    {
      this.InsertChart( chart, paragraph );
    }

    private void InsertChart( Chart chart, Paragraph paragraph )
    {
      Paragraph p;

      // Create a new chart part uri.
      var chartPartUriPath = String.Empty;
      var chartIndex = 1;
      do
      {
        chartPartUriPath = String.Format( "/word/charts/chart{0}.xml", chartIndex );
        chartIndex++;
      } while( _package.PartExists( new Uri( chartPartUriPath, UriKind.Relative ) ) );

      // Create chart part.
      var chartPackagePart = _package.CreatePart( new Uri( chartPartUriPath, UriKind.Relative ), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", CompressionOption.Normal );

      // Create a new chart relationship
      var relID = this.GetNextFreeRelationshipID();
      var rel = this.PackagePart.CreateRelationship( chartPackagePart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", relID );

      // Save a chart info the chartPackagePart
      if( paragraph == null )
      {
        using( TextWriter tw = new StreamWriter( new PackagePartStream( chartPackagePart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
        {
          chart.Xml.Save( tw );
        }
        p = InsertParagraph();
      }
      else
      {
        using( TextWriter tw = new StreamWriter( chartPackagePart.GetStream( FileMode.Create, FileAccess.Write ) ) )
        {
          chart.Xml.Save( tw );
        }
        p = paragraph;
      }

      // Insert a new chart into a paragraph.
      var chartElement = new XElement( XName.Get( "r", w.NamespaceName ),
                                       new XElement( XName.Get( "drawing", w.NamespaceName ),
                                                     new XElement( XName.Get( "inline", wp.NamespaceName ),
                                                                   new XElement( XName.Get( "extent", wp.NamespaceName ), new XAttribute( "cx", "5486400" ), new XAttribute( "cy", "3200400" ) ),
                                                                   new XElement( XName.Get( "effectExtent", wp.NamespaceName ), new XAttribute( "l", "0" ), new XAttribute( "t", "0" ), new XAttribute( "r", "19050" ), new XAttribute( "b", "19050" ) ),
                                                                   new XElement( XName.Get( "docPr", wp.NamespaceName ), new XAttribute( "id", "1" ), new XAttribute( "name", "chart" ) ),
                                                                   new XElement( XName.Get( "graphic", a.NamespaceName ),
                                                                                 new XElement( XName.Get( "graphicData", a.NamespaceName ),
                                                                                               new XAttribute( "uri", c.NamespaceName ),
                                                                                               new XElement( XName.Get( "chart", c.NamespaceName ),
                                                                                                             new XAttribute( XName.Get( "id", r.NamespaceName ), relID ) ) ) ) ) ) );
      p.Xml.Add( chartElement );
    }

    public List<Section> GetSections()
    {
      var paragraphs = Paragraphs;
      var sections = new List<Section>();
      var sectionParagraphs = new List<Paragraph>();

      foreach( Paragraph paragraph in paragraphs )
      {

        var sectionInPara = paragraph.Xml.Descendants().FirstOrDefault( s => s.Name.LocalName == "sectPr" );

        if( sectionInPara != null )
        {
          sectionParagraphs.Add( paragraph );

          var section = new Section( Document, sectionInPara );
          section.SectionParagraphs = sectionParagraphs;

          sections.Add( section );
          sectionParagraphs = new List<Paragraph>();
        }
        else
        {
          sectionParagraphs.Add( paragraph );
        }
      }

      XElement body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
      XElement baseSectionXml = body.Element( XName.Get( "sectPr", w.NamespaceName ) );

      var baseSection = new Section( Document, baseSectionXml );
      baseSection.SectionParagraphs = sectionParagraphs;
      sections.Add( baseSection );

      return sections;
    }

    /// <summary>
    /// Create a new List
    /// </summary>
    public List AddList( string listText = null, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false )
    {
      return AddListItem( new List( this, null ), listText, level, listType, startNumber, trackChanges, continueNumbering );
    }

    /// <summary>
    /// Add a list item to an existing list
    /// </summary>
    public List AddListItem( List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false )
    {
      if( startNumber.HasValue && continueNumbering )
        throw new InvalidOperationException( "Cannot specify a start number and at the same time continue numbering from another list" );

      var result = HelperFunctions.CreateItemInList( list, listText, level, listType, startNumber, trackChanges, continueNumbering );
      var lastItem = result.Items.LastOrDefault();

      if( lastItem != null )
      {
        lastItem.PackagePart = this.PackagePart;
      }

      return result;
    }

    /// <summary>
    /// Insert a list in the document
    /// </summary>
    /// <param name="list">The list to insert into the document.</param>
    /// <returns>The list that was inserted into the document.</returns>
    public override List InsertList( List list )
    {
      base.InsertList( list );
      return list;
    }

    public override List InsertList( List list, Font fontFamily, double fontSize )
    {
      base.InsertList( list, fontFamily, fontSize );
      return list;
    }

    public override List InsertList( List list, double fontSize )
    {
      base.InsertList( list, fontSize );
      return list;
    }

    /// <summary>
    /// Insert a list at an index location in the document
    /// </summary>
    /// <param name="index">Index in document to insert the list.</param>
 	  /// <param name="list">The list that was inserted into the document.</param>
    /// <returns></returns>
    public new List InsertList( int index, List list )
    {
      base.InsertList( index, list );
      return list;
    }

    /// <summary>
    /// Insert a default Table of Contents in the current document
    /// </summary>
    public TableOfContents InsertDefaultTableOfContents()
    {
      return InsertTableOfContents( "Table of contents", TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U );
    }

    /// <summary>
    /// Insert a Table of Contents in the current document
    /// </summary>
    public TableOfContents InsertTableOfContents( string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null )
    {
      var toc = TableOfContents.CreateTableOfContents( this, title, switches, headerStyle, maxIncludeLevel, rightTabPos );
      Xml.Add( toc.Xml );
      return toc;
    }

    /// <summary>
    /// Insert a Table of Contents in the current document at a specific location (prior to the referenced paragraph)
    /// </summary>
    public TableOfContents InsertTableOfContents( Paragraph reference, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null )
    {
      var toc = TableOfContents.CreateTableOfContents( this, title, switches, headerStyle, maxIncludeLevel, rightTabPos );
      reference.Xml.AddBeforeSelf( toc.Xml );
      return toc;
    }

    /// <summary>
    /// Copy the Document into a new Document
    /// </summary>
    /// <returns>Returns a copy of a the Document</returns>
    public DocX Copy()
    {
      var memorystream = new MemoryStream();
      this.SaveAs( memorystream );
      memorystream.Seek( 0, SeekOrigin.Begin );
      return Load( memorystream );
    }

    public void AddPasswordProtection( EditRestrictions editRestrictions, string password )
    {
      // Intellectual Property information :
      //
      // The following code handles password protection of Word documents (Open Specifications) 
      // and is an implementation of algorithm(s) described in Office Document Cryptography Structure 
      // here: https://msdn.microsoft.com/en-us/library/cc313071.aspx.
      //
      // The code’s use is covered under Microsoft’s Open Specification Promise 
      // described here: https://msdn.microsoft.com/en-US/openspecifications/dn646765


      // Remove existing password protection
      this.RemoveProtection();

      // If no EditRestrictions, nothing to do
      if( editRestrictions == EditRestrictions.none )
        return;

      // Variables
      int maxPasswordLength = 15;
      var saltArray = new byte[ 16 ];
      var keyValues = new byte[ 14 ];

      // Init DocumentProtection element
      var documentProtection = new XElement( XName.Get( "documentProtection", w.NamespaceName ) );
      documentProtection.Add( new XAttribute( XName.Get( "edit", w.NamespaceName ), editRestrictions.ToString() ) );
      documentProtection.Add( new XAttribute( XName.Get( "enforcement", w.NamespaceName ), "1" ) );

      int[] InitialCodeArray = { 0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3 };
      int[,] EncryptionMatrix = new int[ 15, 7 ]
      {
            /* char 1  */ { 0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09},
            /* char 2  */ { 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF},
            /* char 3  */ { 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0},
            /* char 4  */ { 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40},
            /* char 5  */ { 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5},
            /* char 6  */ { 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A},
            /* char 7  */ { 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9},
            /* char 8  */ { 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0},
            /* char 9  */ { 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC},
            /* char 10 */ { 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10},
            /* char 11 */ { 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168},
            /* char 12 */ { 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C},
            /* char 13 */ { 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD},
            /* char 14 */ { 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC},
            /* char 15 */ { 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4}
      };

      // Generate the salt
      var random = new RNGCryptoServiceProvider();
      random.GetNonZeroBytes( saltArray );

      // Validate the provided password
      if( !String.IsNullOrEmpty( password ) )
      {
        password = password.Substring( 0, Math.Min( password.Length, maxPasswordLength ) );
        var byteChars = new byte[ password.Length ];

        for( int i = 0; i < password.Length; i++ )
        {
          var temp = Convert.ToInt32( password[ i ] );
          byteChars[ i ] = Convert.ToByte( temp & 0x00FF );

          if( byteChars[ i ] == 0 )
          {
            byteChars[ i ] = Convert.ToByte( ( temp & 0x00FF ) >> 8 );
          }
        }

        var intHighOrderWord = InitialCodeArray[ byteChars.Length - 1 ];

        for( int i = 0; i < byteChars.Length; i++ )
        {
          int tmp = maxPasswordLength - byteChars.Length + i;
          for( int intBit = 0; intBit < 7; intBit++ )
          {
            if( ( byteChars[ i ] & ( 0x0001 << intBit ) ) != 0 )
            {
              intHighOrderWord ^= EncryptionMatrix[ tmp, intBit ];
            }
          }
        }

        int intLowOrderWord = 0;

        // For each character in the strPassword, going backwards
        for( int i = byteChars.Length - 1; i >= 0; i-- )
        {
          intLowOrderWord = ( ( ( intLowOrderWord >> 14 ) & 0x0001 ) | ( ( intLowOrderWord << 1 ) & 0x7FFF ) ) ^ byteChars[ i ];
        }

        intLowOrderWord = ( ( ( intLowOrderWord >> 14 ) & 0x0001 ) | ( ( intLowOrderWord << 1 ) & 0x7FFF ) ) ^ byteChars.Length ^ 0xCE4B;

        // Combine the Low and High Order Word
        var intCombinedkey = ( intHighOrderWord << 16 ) + intLowOrderWord;

        // The byte order of the result shall be reversed [Example: 0x64CEED7E becomes 7EEDCE64. end example],
        // and that value shall be hashed as defined by the attribute values.

        for( int i = 0; i < 4; i++ )
        {
          keyValues[ i ] = Convert.ToByte( ( ( uint )( intCombinedkey & ( 0x000000FF << ( i * 8 ) ) ) ) >> ( i * 8 ) );
        }
      }

      var sb = new StringBuilder();
      for( int intTemp = 0; intTemp < 4; intTemp++ )
      {
        sb.Append( Convert.ToString( keyValues[ intTemp ], 16 ) );
      }

      keyValues = Encoding.Unicode.GetBytes( sb.ToString().ToUpper() );
      keyValues = MergeArrays( keyValues, saltArray );

      int iterations = 100000;

      var sha1 = new SHA1Managed();
      keyValues = sha1.ComputeHash( keyValues );
      var iterator = new byte[ 4 ];
      for( int i = 0; i < iterations; i++ )
      {
        iterator[ 0 ] = Convert.ToByte( ( i & 0x000000FF ) >> 0 );
        iterator[ 1 ] = Convert.ToByte( ( i & 0x0000FF00 ) >> 8 );
        iterator[ 2 ] = Convert.ToByte( ( i & 0x00FF0000 ) >> 16 );
        iterator[ 3 ] = Convert.ToByte( ( i & 0xFF000000 ) >> 24 );

        keyValues = MergeArrays( iterator, keyValues );
        keyValues = sha1.ComputeHash( keyValues );
      }

      documentProtection.Add( new XAttribute( XName.Get( "cryptProviderType", w.NamespaceName ), "rsaFull" ) );
      documentProtection.Add( new XAttribute( XName.Get( "cryptAlgorithmClass", w.NamespaceName ), "hash" ) );
      documentProtection.Add( new XAttribute( XName.Get( "cryptAlgorithmType", w.NamespaceName ), "typeAny" ) );
      documentProtection.Add( new XAttribute( XName.Get( "cryptAlgorithmSid", w.NamespaceName ), "4" ) );
      documentProtection.Add( new XAttribute( XName.Get( "cryptSpinCount", w.NamespaceName ), iterations.ToString() ) );
      documentProtection.Add( new XAttribute( XName.Get( "hash", w.NamespaceName ), Convert.ToBase64String( keyValues ) ) );
      documentProtection.Add( new XAttribute( XName.Get( "salt", w.NamespaceName ), Convert.ToBase64String( saltArray ) ) );

      _settings.Root.AddFirst( documentProtection );
    }

    #endregion

    #region Internal Methods

    internal float getMarginAttribute( XName name )
    {
      var body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
      var sectPr = body.Element( XName.Get( "sectPr", w.NamespaceName ) );
      var pgMar = sectPr?.Element( XName.Get( "pgMar", w.NamespaceName ) );
      var top = pgMar?.Attribute( name );
      if( top != null )
      {
        float f;
        if( float.TryParse( top.Value, out f ) )
          return ( int )( f / _pageSizeMultiplier );
      }

      return 0;
    }

    internal void setMarginAttribute( XName xName, float value )
    {
      var body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
      var sectPr = body.Element( XName.Get( "sectPr", w.NamespaceName ) );
      var pgMar = sectPr?.Element( XName.Get( "pgMar", w.NamespaceName ) );
      var top = pgMar?.Attribute( xName );
      top?.SetValue( value * Convert.ToInt32( _pageSizeMultiplier ) );
    }

    internal string GetCollectiveText( List<PackagePart> list )
    {
      var text = string.Empty;

      foreach( var hp in list )
      {
        using( TextReader tr = new StreamReader( hp.GetStream() ) )
        {
          var d = XDocument.Load( tr );

          var sb = new StringBuilder();

          // Loop through each text item in this run
          foreach( XElement descendant in d.Descendants() )
          {
            switch( descendant.Name.LocalName )
            {
              case "tab":
                sb.Append( "\t" );
                break;
              case "br":
                sb.Append( "\n" );
                break;
              case "t":
                goto case "delText";
              case "delText":
                sb.Append( descendant.Value );
                break;
              default:
                break;
            }
          }

          text += "\n" + sb;
        }
      }

      return text;
    }

    internal static void PostCreation( Package package, DocumentTypes documentType = DocumentTypes.Document )
    {
      XDocument mainDoc, stylesDoc, numberingDoc;

      #region MainDocumentPart
      // Create the main document part for this package
      var mainDocPart = ( documentType == DocumentTypes.Document )
                        ? package.CreatePart( new Uri( "/word/document.xml", UriKind.Relative ), HelperFunctions.DOCUMENT_DOCUMENTTYPE, CompressionOption.Normal )
                        : package.CreatePart( new Uri( "/word/document.xml", UriKind.Relative ), HelperFunctions.TEMPLATE_DOCUMENTTYPE, CompressionOption.Normal );

      package.CreateRelationship( mainDocPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" );

      // Load the document part into a XDocument object
      using( TextReader tr = new StreamReader( mainDocPart.GetStream( FileMode.Create, FileAccess.ReadWrite ) ) )
      {
        mainDoc = XDocument.Parse
        ( @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                   <w:document xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                   <w:body>
                    <w:sectPr w:rsidR=""003E25F4"" w:rsidSect=""00FC3028"">
                        <w:pgSz w:w=""11906"" w:h=""16838""/>
                        <w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""708"" w:footer=""708"" w:gutter=""0""/>
                        <w:cols w:space=""708""/>
                        <w:docGrid w:linePitch=""360""/>
                    </w:sectPr>
                   </w:body>
                   </w:document>"
        );
      }

      // Save the main document
      using( TextWriter tw = new StreamWriter( new PackagePartStream( mainDocPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        mainDoc.Save( tw, SaveOptions.None );
      }
      #endregion

      #region StylePart
      stylesDoc = HelperFunctions.AddDefaultStylesXml( package );
      #endregion

      #region NumberingPart
      numberingDoc = HelperFunctions.AddDefaultNumberingXml( package );
      #endregion

      package.Close();
    }

    internal static DocX PostLoad( ref Package package )
    {
      var document = new DocX( null, null );
      document._package = package;
      document.Document = document;

      #region MainDocumentPart
      document.PackagePart = HelperFunctions.GetMainDocumentPart( package );

      using( TextReader tr = new StreamReader( document.PackagePart.GetStream( FileMode.Open, FileAccess.Read ) ) )
      {
        document._mainDoc = XDocument.Load( tr, LoadOptions.PreserveWhitespace );
      }
      #endregion

      DocX.PopulateDocument( document, package );

      using( TextReader tr = new StreamReader( document._settingsPart.GetStream() ) )
      {
        document._settings = XDocument.Load( tr );
      }

      document._paragraphLookup.Clear();
      var paragraphs = document.Paragraphs;
      foreach( var p in paragraphs )
      {
        if( !document._paragraphLookup.ContainsKey( p._endIndex ) )
        {
          document._paragraphLookup.Add( p._endIndex, p );
        }
      }

      return document;
    }

    internal void AddHyperlinkStyleIfNotPresent()
    {
      var word_styles_Uri = new Uri( "/word/styles.xml", UriKind.Relative );

      // If the internal document contains no /word/styles.xml create one.
      if( !_package.PartExists( word_styles_Uri ) )
      {
        HelperFunctions.AddDefaultStylesXml( _package );
      }

      // Load the styles.xml into memory.
      XDocument word_styles;
      using( TextReader tr = new StreamReader( _package.GetPart( word_styles_Uri ).GetStream() ) )
      {
        word_styles = XDocument.Load( tr );
      }

      bool hyperlinkStyleExists =
      (
          from s in word_styles.Element( w + "styles" ).Elements()
          let styleId = s.Attribute( XName.Get( "styleId", w.NamespaceName ) )
          where ( styleId != null && styleId.Value == "Hyperlink" )
          select s
      ).Count() > 0;

      if( !hyperlinkStyleExists )
      {
        var style = new XElement
        (
            w + "style",
            new XAttribute( w + "type", "character" ),
            new XAttribute( w + "styleId", "Hyperlink" ),
                new XElement( w + "name", new XAttribute( w + "val", "Hyperlink" ) ),
                new XElement( w + "basedOn", new XAttribute( w + "val", "DefaultParagraphFont" ) ),
                new XElement( w + "uiPriority", new XAttribute( w + "val", "99" ) ),
                new XElement( w + "unhideWhenUsed" ),
                new XElement( w + "rsid", new XAttribute( w + "val", "0005416C" ) ),
                new XElement
                (
                    w + "rPr",
                    new XElement( w + "color", new XAttribute( w + "val", "0000FF" ), new XAttribute( w + "themeColor", "hyperlink" ) ),
                    new XElement
                    (
                        w + "u",
                        new XAttribute( w + "val", "single" )
                    )
                )
        );
        word_styles.Element( w + "styles" ).Add( style );

        // Save the styles document.
        using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( word_styles_Uri ).GetStream() ) ) )
        {
          word_styles.Save( tw );
        }
      }
    }

    /// <summary>
    /// Adds a Header to a document.
    /// If the document already contains a Header it will be replaced.
    /// </summary>
    /// <returns>The Header that was added to the document.</returns>
    internal void AddHeadersOrFooters( bool b )
    {
      var element = b ? "hdr" : "ftr";
      var reference = b ? "header" : "footer";

      this.DeleteHeadersOrFooters( b );

      var sectPr = _mainDoc.Root.Element( w + "body" ).Element( w + "sectPr" );

      for( int i = 1; i < 4; i++ )
      {
        var header_uri = string.Format( "/word/{0}{1}.xml", reference, i );

        var headerPart = _package.CreatePart( new Uri( header_uri, UriKind.Relative ), string.Format( "application/vnd.openxmlformats-officedocument.wordprocessingml.{0}+xml", reference ), CompressionOption.Normal );
        var headerRelationship = this.PackagePart.CreateRelationship( headerPart.Uri, TargetMode.Internal, string.Format( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference ) );

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
        switch( i )
        {
          case 1:
            type = "default";
            break;
          case 2:
            type = "even";
            break;
          case 3:
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
    }

    internal void DeleteHeadersOrFooters( bool b )
    {
      string reference = "footer";
      if( b )
        reference = "header";

      // Get all header Relationships in this document.
      var header_relationships = this.PackagePart.GetRelationshipsByType( string.Format( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference ) );

      foreach( PackageRelationship header_relationship in header_relationships )
      {
        // Get the TargetUri for this Part.
        Uri header_uri = header_relationship.TargetUri;

        // Check to see if the document actually contains the Part.
        if( !header_uri.OriginalString.StartsWith( "/word/" ) )
          header_uri = new Uri( "/word/" + header_uri.OriginalString, UriKind.Relative );

        if( _package.PartExists( header_uri ) )
        {
          // Delete the Part
          _package.DeletePart( header_uri );

          // Get all references to this Relationship in the document.
          var query =
          (
              from e in _mainDoc.Descendants( XName.Get( "body", w.NamespaceName ) ).Descendants()
              where ( e.Name.LocalName == string.Format( "{0}Reference", reference ) ) && ( e.Attribute( r + "id" ).Value == header_relationship.Id )
              select e
          );

          // Remove all references to this Relationship in the document.
          for( int i = 0; i < query.Count(); i++ )
            query.ElementAt( i ).Remove();

          // Delete the Relationship.
          _package.DeleteRelationship( header_relationship.Id );
        }
      }
    }

    internal Image AddImage( object o, string contentType = "image/jpeg" )
    {
      // Open a Stream to the new image being added.
      var newImageStream = ( o is string ) ? new FileStream( o as string, FileMode.Open, FileAccess.Read ) : o as Stream;

      // Get all image parts in word\document.xml
      PackagePartCollection packagePartCollection = _package.GetParts();
      var parts = packagePartCollection.Select( x => new
      {
        UriString = x.Uri.ToString(),
        Part = x
      } ).ToList();

      var partLookup = parts.ToDictionary( x => x.UriString, x => x.Part, StringComparer.Ordinal );

      List<PackagePart> imageParts = new List<PackagePart>();
      foreach( var item in this.PackagePart.GetRelationshipsByType( RelationshipImage ) )
      {
        var targetUri = item.TargetUri.ToString();
        PackagePart part;
        if( partLookup.TryGetValue( targetUri, out part ) )
        {
          imageParts.Add( part );
        }
      }

      IEnumerable<PackagePart> relsParts = parts
        .Where(
          part =>
          part.Part.ContentType.Equals( ContentTypeApplicationRelationShipXml, StringComparison.Ordinal ) &&
          part.UriString.IndexOf( "/word/", StringComparison.Ordinal ) > -1 )
        .Select( part => part.Part );

      XName xNameTarget = XName.Get( "Target" );
      XName xNameTargetMode = XName.Get( "TargetMode" );

      foreach( PackagePart relsPart in relsParts )
      {
        XDocument relsPartContent;
        using( TextReader tr = new StreamReader( relsPart.GetStream( FileMode.Open, FileAccess.Read ) ) )
        {
          relsPartContent = XDocument.Load( tr );
        }

        var imageRelationships =
        relsPartContent.Root.Elements().Where
        (
            imageRel =>
            imageRel.Attribute( XName.Get( "Type" ) ).Value.Equals( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" )
        );

        foreach( XElement imageRelationship in imageRelationships )
        {
          XAttribute attribute = imageRelationship.Attribute( xNameTarget );
          if( attribute != null )
          {
            var targetModeAttr = imageRelationship.Attribute( xNameTargetMode );
            var targetMode = ( targetModeAttr != null ) ? targetModeAttr.Value : string.Empty;

            if( !targetMode.Equals( "External" ) )
            {
              var imagePartUri = Path.Combine( Path.GetDirectoryName( relsPart.Uri.ToString() ), attribute.Value );
              imagePartUri = Path.GetFullPath( imagePartUri.Replace( "\\_rels", string.Empty ) );
              imagePartUri = imagePartUri.Replace( Path.GetFullPath( "\\" ), string.Empty ).Replace( "\\", "/" );

              if( !imagePartUri.StartsWith( "/" ) )
              {
                imagePartUri = "/" + imagePartUri;
              }

              var imagePart = _package.GetPart( new Uri( imagePartUri, UriKind.Relative ) );
              imageParts.Add( imagePart );
            }
          }
        }
      }

      // Loop through each image part in this document.
      foreach( PackagePart pp in imageParts )
      {
        // Get the image object for this image part.
        using( Stream tempStream = pp.GetStream( FileMode.Open, FileAccess.Read ) )
        {
          // Compare this image to the new image being added.
          if( HelperFunctions.IsSameFile( tempStream, newImageStream ) )
          {
            // Return the Image object
            PackageRelationship relationship = this.PackagePart.GetRelationshipsByType( RelationshipImage ).First( x => x.TargetUri == pp.Uri );
            return new Image( this, relationship );
          }
        }
      }

      var imgPartUriPath = string.Empty;
      var extension = contentType.Substring( contentType.LastIndexOf( "/" ) + 1 );
      do
      {
        // Create a new image part.
        imgPartUriPath = string.Format
        (
            "/word/media/{0}.{1}",
            Guid.NewGuid(), // The unique part.
            extension
        );

      } while( _package.PartExists( new Uri( imgPartUriPath, UriKind.Relative ) ) );

      // We are now guaranteed that imgPartUriPath is unique.
      var img = _package.CreatePart( new Uri( imgPartUriPath, UriKind.Relative ), contentType, CompressionOption.Normal );

      // Create a new image relationship
      var rel = this.PackagePart.CreateRelationship( img.Uri, TargetMode.Internal, RelationshipImage );

      // Open a Stream to the newly created Image part.
      using( Stream stream = new PackagePartStream( img.GetStream( FileMode.Create, FileAccess.Write ) ) )
      {
        // Using the Stream to the real image, copy this streams data into the newly create Image part.
        using( newImageStream )
        {
          CopyStream( newImageStream, stream, bufferSize: 4096 );
        }// Close the Stream to the new image.
      }// Close the Stream to the new image part.

      return new Image( this, rel );
    }

    internal static void UpdateCorePropertyValue( DocX document, string corePropertyName, string corePropertyValue )
    {
      var matchPattern = string.Format( @"(DOCPROPERTY)?{0}\\\*MERGEFORMAT", corePropertyName ).ToLower();
      foreach( XElement e in document._mainDoc.Descendants( XName.Get( "fldSimple", w.NamespaceName ) ) )
      {
        var attr_value = e.Attribute( XName.Get( "instr", w.NamespaceName ) ).Value.Replace( " ", string.Empty ).Trim().ToLower();

        if( Regex.IsMatch( attr_value, matchPattern ) )
        {
          var firstRun = e.Element( w + "r" );
          var firstText = firstRun.Element( w + "t" );
          var rPr = firstText.Element( w + "rPr" );

          // Delete everything and insert updated text value
          e.RemoveNodes();

          var t = new XElement( w + "t", rPr, corePropertyValue );
          Xceed.Words.NET.Text.PreserveSpace( t );
          e.Add( new XElement( firstRun.Name, firstRun.Attributes(), firstRun.Element( XName.Get( "rPr", w.NamespaceName ) ), t ) );
        }
      }

      #region Headers

      var headerParts = from headerPart in document._package.GetParts()
                        where ( Regex.IsMatch( headerPart.Uri.ToString(), @"/word/header\d?.xml" ) )
                        select headerPart;
      foreach( PackagePart pp in headerParts )
      {
        var header = XDocument.Load( new StreamReader( pp.GetStream() ) );

        foreach( XElement e in header.Descendants( XName.Get( "fldSimple", w.NamespaceName ) ) )
        {
          string attr_value = e.Attribute( XName.Get( "instr", w.NamespaceName ) ).Value.Replace( " ", string.Empty ).Trim().ToLower();
          if( Regex.IsMatch( attr_value, matchPattern ) )
          {
            var firstRun = e.Element( w + "r" );

            // Delete everything and insert updated text value
            e.RemoveNodes();

            var t = new XElement( w + "t", corePropertyValue );
            Xceed.Words.NET.Text.PreserveSpace( t );
            e.Add( new XElement( firstRun.Name, firstRun.Attributes(), firstRun.Element( XName.Get( "rPr", w.NamespaceName ) ), t ) );
          }
        }

        using( TextWriter tw = new StreamWriter( new PackagePartStream( pp.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
        {
          header.Save( tw );
        }
      }
      #endregion

      #region Footers
      var footerParts = from footerPart in document._package.GetParts()
                        where ( Regex.IsMatch( footerPart.Uri.ToString(), @"/word/footer\d?.xml" ) )
                        select footerPart;
      foreach( PackagePart pp in footerParts )
      {
        var footer = XDocument.Load( new StreamReader( pp.GetStream() ) );

        foreach( XElement e in footer.Descendants( XName.Get( "fldSimple", w.NamespaceName ) ) )
        {
          string attr_value = e.Attribute( XName.Get( "instr", w.NamespaceName ) ).Value.Replace( " ", string.Empty ).Trim().ToLower();
          if( Regex.IsMatch( attr_value, matchPattern ) )
          {
            var firstRun = e.Element( w + "r" );

            // Delete everything and insert updated text value
            e.RemoveNodes();

            var t = new XElement( w + "t", corePropertyValue );
            Xceed.Words.NET.Text.PreserveSpace( t );
            e.Add( new XElement( firstRun.Name, firstRun.Attributes(), firstRun.Element( XName.Get( "rPr", w.NamespaceName ) ), t ) );
          }
        }

        using( TextWriter tw = new StreamWriter( new PackagePartStream( pp.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
        {
          footer.Save( tw );
        }
      }
      #endregion
      DocX.PopulateDocument( document, document._package );
    }

    /// <summary>
    /// Update the custom properties inside the document
    /// </summary>
    /// <param name="document">The DocX document</param>
    /// <param name="customPropertyName">The property used inside the document</param>
    /// <param name="customPropertyValue">The new value for the property</param>
    /// <remarks>Different version of Word create different Document XML.</remarks>
    internal static void UpdateCustomPropertyValue( DocX document, string customPropertyName, string customPropertyValue )
    {
      // A list of documents, which will contain, The Main Document and if they exist: header1, header2, header3, footer1, footer2, footer3.
      var documents = new List<XElement> { document._mainDoc.Root };

      // Check if each header exists and add if if so.
      #region Headers
      var headers = document.Headers;
      if( headers.First != null )
      {
        documents.Add( headers.First.Xml );
      }
      if( headers.Odd != null )
      {
        documents.Add( headers.Odd.Xml );
      }
      if( headers.Even != null )
      {
        documents.Add( headers.Even.Xml );
      }
      #endregion

      // Check if each footer exists and add if if so.
      #region Footers
      var footers = document.Footers;
      if( footers.First != null )
      {
        documents.Add( footers.First.Xml );
      }
      if( footers.Odd != null )
      {
        documents.Add( footers.Odd.Xml );
      }
      if( footers.Even != null )
      {
        documents.Add( footers.Even.Xml );
      }
      #endregion

      var match_value = string.Format( @"DOCPROPERTY  {0}  \* MERGEFORMAT", customPropertyName.Contains( " " ) ? "\"" + customPropertyName + "\"" : customPropertyName )
                              .Replace( " ", string.Empty );

      // Process each document in the list.
      foreach( XElement doc in documents )
      {
        #region Word 2010+
        foreach( XElement e in doc.Descendants( XName.Get( "instrText", w.NamespaceName ) ) )
        {
          var attr_value = e.Value.Replace( " ", string.Empty ).Trim();          

          if( attr_value.Equals( match_value, StringComparison.CurrentCultureIgnoreCase ) )
          {
            var node = e.Parent.NextNode;
            bool found = false;
            while( true )
            {
              if( node.NodeType == XmlNodeType.Element )
              {
                var ele = node as XElement;
                var match = ele.Descendants( XName.Get( "t", w.NamespaceName ) );
                if( match.Any() )
                {
                  if( !found )
                  {
                    match.First().Value = customPropertyValue;
                    found = true;
                  }
                  else
                  {
                    ele.RemoveNodes();
                  }
                }
                else
                {
                  match = ele.Descendants( XName.Get( "fldChar", w.NamespaceName ) );
                  if( match.Any() )
                  {
                    var endMatch = match.First().Attribute( XName.Get( "fldCharType", w.NamespaceName ) );
                    if( endMatch != null && endMatch.Value == "end" )
                    {
                      break;
                    }
                  }
                }
              }
              node = node.NextNode;
            }
          }
        }
        #endregion

        #region < Word 2010
        foreach( XElement e in doc.Descendants( XName.Get( "fldSimple", w.NamespaceName ) ) )
        {
          var attr_value = e.Attribute( XName.Get( "instr", w.NamespaceName ) ).Value.Replace( " ", string.Empty ).Trim();

          if( attr_value.Equals( match_value, StringComparison.CurrentCultureIgnoreCase ) )
          {
            var firstRun = e.Element( w + "r" );
            var firstText = firstRun.Element( w + "t" );
            var rPr = firstText.Element( w + "rPr" );

            // Delete everything and insert updated text value
            e.RemoveNodes();

            var t = new XElement( w + "t", rPr, customPropertyValue );
            Xceed.Words.NET.Text.PreserveSpace( t );
            e.Add( new XElement( firstRun.Name, firstRun.Attributes(), firstRun.Element( XName.Get( "rPr", w.NamespaceName ) ), t ) );
          }
        }
        #endregion
      }
    }

    internal XDocument AddStylesForList()
    {
      var fileUri = new Uri( "/word/styles.xml", UriKind.Relative );

      // Create the /word/styles.xml if it doesn't already exist
      if( !_package.PartExists( fileUri ) )
        HelperFunctions.AddDefaultStylesXml( _package );

      // Load the xml into memory.
      XDocument stylesDoc;
      using( TextReader tr = new StreamReader( _package.GetPart( fileUri ).GetStream() ) )
        stylesDoc = XDocument.Load( tr );

      bool listStyleExists =
      (
          from s in stylesDoc.Element( w + "styles" ).Elements()
          let styleId = s.Attribute( XName.Get( "styleId", w.NamespaceName ) )
          where ( styleId != null && styleId.Value == "ListParagraph" )
          select s
      ).Any();

      if( !listStyleExists )
      {
        var style = new XElement
        (
            w + "style",
            new XAttribute( w + "type", "paragraph" ),
            new XAttribute( w + "styleId", "ListParagraph" ),
                new XElement( w + "name", new XAttribute( w + "val", "List Paragraph" ) ),
                new XElement( w + "basedOn", new XAttribute( w + "val", "Normal" ) ),
                new XElement( w + "uiPriority", new XAttribute( w + "val", "34" ) ),
                new XElement( w + "qformat" ),
                new XElement( w + "rsid", new XAttribute( w + "val", "00832EE1" ) ),
                new XElement
                (
                    w + "rPr",
                    new XElement( w + "ind", new XAttribute( w + "left", "720" ) ),
                    new XElement
                    (
                        w + "contextualSpacing"
                    )
                )
        );
        stylesDoc.Element( w + "styles" ).Add( style );

        // Save the document
        using( TextWriter tw = new StreamWriter( _package.GetPart( fileUri ).GetStream() ) )
          stylesDoc.Save( tw );
      }

      return stylesDoc;
    }

    internal bool getMirrorMargins(XName name)
    {
      var body = _mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));
      var sectPr = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));
      var mirrorMargins = sectPr?.Element(XName.Get("mirrorMargins", DocX.w.NamespaceName));
      return (mirrorMargins != null);
    }

    internal void setMirrorMargins(XName name, bool value)
    {
      var body = _mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));
      var sectPr = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));
      var mirrorMargins = sectPr?.Element(XName.Get("mirrorMargins", DocX.w.NamespaceName));

      if (mirrorMargins == null)
        sectPr.Add(new XElement(w + "mirrorMargins", string.Empty));

      else
      {
        if (!value)
          mirrorMargins.Remove();
      }
    }

    /// <summary>
    /// Finds the next free Id for bookmarkStart/docPr.
    /// </summary>
    internal long GetNextFreeDocPrId()
    {
      lock( nextFreeDocPrIdLock )
      {
        if( nextFreeDocPrId != null )
        {
          nextFreeDocPrId++;
          return nextFreeDocPrId.Value;
        }

        var xNameBookmarkStart = XName.Get( "bookmarkStart", DocX.w.NamespaceName );
        var xNameDocPr = XName.Get( "docPr", DocX.wp.NamespaceName );

        long newDocPrId = 1;
        HashSet<string> existingIds = new HashSet<string>();
        foreach( var bookmarkId in Xml.Descendants() )
        {
          if( bookmarkId.Name != xNameBookmarkStart && bookmarkId.Name != xNameDocPr )
            continue;

          var idAtt = bookmarkId.Attributes().FirstOrDefault( x => x.Name.LocalName == "id" );
          if( idAtt != null )
            existingIds.Add( idAtt.Value );
        }

        while( existingIds.Contains( newDocPrId.ToString() ) )
          newDocPrId++;

        nextFreeDocPrId = newDocPrId;
        return nextFreeDocPrId.Value;
      }
    }

    #endregion

    #region Private Methods

    private static DocX CreateDocument( DocumentTypes documentType )
    {
      // Store this document in memory
      var ms = new MemoryStream();

      // Create the docx package
      var package = Package.Open( ms, FileMode.Create, FileAccess.ReadWrite );

      DocX.PostCreation( package, documentType );
      return DocX.Load( ms );
    }

    private Header GetHeaderByType( string type )
    {
      return ( Header )GetHeaderOrFooterByType( type, true );
    }

    private Footer GetFooterByType( string type )
    {
      return ( Footer )GetHeaderOrFooterByType( type, false );
    }

    private object GetHeaderOrFooterByType( string type, bool isHeader )
    {
      // Switch which handles either case Header\Footer, this just cuts down on code duplication.
      string reference = "footerReference";
      if( isHeader )
        reference = "headerReference";

      // Get the Id of the [default, even or first] [Header or Footer]
      string Id =
      (
          from e in _mainDoc.Descendants( XName.Get( "body", w.NamespaceName ) ).Descendants()
          where ( e.Name.LocalName == reference ) && ( e.Attribute( w + "type" ).Value == type )
          select e.Attribute( r + "id" ).Value
      ).LastOrDefault();

      if( Id != null )
      {
        // Get the Xml file for this Header or Footer.
        var partUri = this.PackagePart.GetRelationship( Id ).TargetUri;

        // Weird problem with PackaePart API.
        if( !partUri.OriginalString.StartsWith( "/word/" ) )
          partUri = new Uri( "/word/" + partUri.OriginalString, UriKind.Relative );

        // Get the Part and open a stream to get the Xml file.
        var part = _package.GetPart( partUri );

        using( TextReader tr = new StreamReader( part.GetStream() ) )
        {
          var doc = XDocument.Load( tr );

          // Header and Footer extend Container.
          Container c;
          if( isHeader )
          {
            c = new Header( this, doc.Element( w + "hdr" ), part );
          }
          else
          {
            c = new Footer( this, doc.Element( w + "ftr" ), part );
          }

          return c;
        }
      }

      // If we got this far something went wrong.
      return null;
    }

    private void merge_images( PackagePart remote_pp, DocX remote_document, XDocument remote_mainDoc, String contentType )
    {
      // Before doing any other work, check to see if this image is actually referenced in the document.
      // In my testing I have found cases of Images inside documents that are not referenced
      var remote_rel = remote_document.PackagePart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( remote_pp.Uri.OriginalString.Replace( "/word/", "" ) ) ).FirstOrDefault();
      if( remote_rel == null )
      {
        remote_rel = remote_document.PackagePart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( remote_pp.Uri.OriginalString ) ).FirstOrDefault();
        if( remote_rel == null )
          return;
      }

      var remote_Id = remote_rel.Id;

      var remote_hash = this.ComputeMD5HashString( remote_pp.GetStream() );
      var image_parts = _package.GetParts().Where( pp => pp.ContentType.Equals( contentType ) );

      bool found = false;
      foreach( var part in image_parts )
      {
        var local_hash = ComputeMD5HashString( part.GetStream() );
        if( local_hash.Equals( remote_hash ) )
        {
          // This image already exists in this document.
          found = true;

          var local_rel = this.PackagePart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( part.Uri.OriginalString.Replace( "/word/", "" ) ) ).FirstOrDefault();

          if( local_rel == null )
          {
            local_rel = this.PackagePart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( part.Uri.OriginalString ) ).FirstOrDefault();
          }

          if( local_rel != null )
          {
            var new_Id = local_rel.Id;

            // Replace all instances of remote_Id in the local document with local_Id
            this.ReplaceAllRemoteID( remote_mainDoc, "blip", "embed", a.NamespaceName, remote_Id, new_Id );
            // Replace all instances of remote_Id in the local document with local_Id (for shapes)
            this.ReplaceAllRemoteID( remote_mainDoc, "imagedata", "id", v.NamespaceName, remote_Id, new_Id );
          }

          break;
        }
      }

      // This image does not exist in this document.
      if( !found )
      {
        var new_uri = remote_pp.Uri.OriginalString;
        new_uri = new_uri.Remove( new_uri.LastIndexOf( "/" ) );
        new_uri += "/" + Guid.NewGuid() + contentType.Replace( "image/", "." );
        if( !new_uri.StartsWith( "/" ) )
        {
          new_uri = "/" + new_uri;
        }

        var new_pp = _package.CreatePart( new Uri( new_uri, UriKind.Relative ), remote_pp.ContentType, CompressionOption.Normal );

        using( Stream s_read = remote_pp.GetStream() )
        {
          using( Stream s_write = new PackagePartStream( new_pp.GetStream( FileMode.Create ) ) )
          {
            CopyStream( s_read, s_write );
          }
        }

        var pr = this.PackagePart.CreateRelationship( new Uri( new_uri, UriKind.Relative ), TargetMode.Internal, RelationshipImage );

        var new_Id = pr.Id;

        //Check if the remote relationship id is a default rId from Word 
        Match relationshipID = Regex.Match( remote_Id, @"rId\d+", RegexOptions.IgnoreCase );

        // Replace all instances of remote_Id in the local document with local_Id
        this.ReplaceAllRemoteID( remote_mainDoc, "blip", "embed", a.NamespaceName, remote_Id, new_Id );

        if( !relationshipID.Success )
        {
          // Replace all instances of remote_Id in the local document with local_Id
          this.ReplaceAllRemoteID( _mainDoc, "blip", "embed", a.NamespaceName, remote_Id, new_Id );

          // Replace all instances of remote_Id in the local document with local_Id (for shapes)
          this.ReplaceAllRemoteID( _mainDoc, "imagedata", "id", v.NamespaceName, remote_Id, new_Id );
        }

        // Replace all instances of remote_Id in the local document with local_Id (for shapes)
        this.ReplaceAllRemoteID( remote_mainDoc, "imagedata", "id", v.NamespaceName, remote_Id, new_Id );
      }
    }

    private void ReplaceAllRemoteID( XDocument remote_mainDoc, string localName, string localNameAttribute, string namespaceName, string remote_Id, string new_Id )
    {
      // Replace all instances of remote_Id in the local document with local_Id
      var elems = remote_mainDoc.Descendants( XName.Get( localName, namespaceName ) );
      foreach( var elem in elems )
      {
        var attribute = elem.Attribute( XName.Get( localNameAttribute, DocX.r.NamespaceName ) );
        if( attribute != null && attribute.Value == remote_Id )
        {
          attribute.SetValue( new_Id );
        }
      }
    }

    private string ComputeMD5HashString( Stream stream )
    {
      MD5 md5 = MD5.Create();
      byte[] hash = md5.ComputeHash( stream );
      StringBuilder sb = new StringBuilder();
      foreach( byte b in hash )
        sb.Append( b.ToString( "X2" ) );
      return sb.ToString();
    }

    private void merge_endnotes( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_endnotes )
    {
      IEnumerable<int> ids =
      (
          from d in _endnotes.Root.Descendants()
          where d.Name.LocalName == "endnote"
          select int.Parse( d.Attribute( XName.Get( "id", w.NamespaceName ) ).Value )
      );

      int max_id = ids.Max() + 1;
      var endnoteReferences = remote_mainDoc.Descendants( XName.Get( "endnoteReference", w.NamespaceName ) );

      foreach( var endnote in remote_endnotes.Root.Elements().OrderBy( fr => fr.Attribute( XName.Get( "id", r.NamespaceName ) ) ).Reverse() )
      {
        XAttribute id = endnote.Attribute( XName.Get( "id", w.NamespaceName ) );
        int i;
        if( id != null && int.TryParse( id.Value, out i ) )
        {
          if( i > 0 )
          {
            foreach( var endnoteRef in endnoteReferences )
            {
              XAttribute a = endnoteRef.Attribute( XName.Get( "id", w.NamespaceName ) );
              if( a != null && int.Parse( a.Value ).Equals( i ) )
              {
                a.SetValue( max_id );
              }
            }

            // We care about copying this footnote.
            endnote.SetAttributeValue( XName.Get( "id", w.NamespaceName ), max_id );
            _endnotes.Root.Add( endnote );
            max_id++;
          }
        }
      }
    }

    private void merge_footnotes( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_footnotes )
    {
      IEnumerable<int> ids =
      (
          from d in _footnotes.Root.Descendants()
          where d.Name.LocalName == "footnote"
          select int.Parse( d.Attribute( XName.Get( "id", DocX.w.NamespaceName ) ).Value )
      );

      int max_id = ids.Max() + 1;
      var footnoteReferences = remote_mainDoc.Descendants( XName.Get( "footnoteReference", DocX.w.NamespaceName ) );

      foreach( var footnote in remote_footnotes.Root.Elements().OrderBy( fr => fr.Attribute( XName.Get( "id", DocX.r.NamespaceName ) ) ).Reverse() )
      {
        XAttribute id = footnote.Attribute( XName.Get( "id", DocX.w.NamespaceName ) );
        int i;
        if( id != null && int.TryParse( id.Value, out i ) )
        {
          if( i > 0 )
          {
            foreach( var footnoteRef in footnoteReferences )
            {
              XAttribute a = footnoteRef.Attribute( XName.Get( "id", DocX.w.NamespaceName ) );
              if( a != null && int.Parse( a.Value ).Equals( i ) )
              {
                a.SetValue( max_id );
              }
            }

            // We care about copying this footnote.
            footnote.SetAttributeValue( XName.Get( "id", DocX.w.NamespaceName ), max_id );
            _footnotes.Root.Add( footnote );
            max_id++;
          }
        }
      }
    }

    private void merge_customs( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc )
    {
      // Get the remote documents custom.xml file.
      XDocument remote_custom_document;
      using( TextReader tr = new StreamReader( remote_pp.GetStream() ) )
      {
        remote_custom_document = XDocument.Load( tr );
      }

      // Get the local documents custom.xml file.
      XDocument local_custom_document;
      using( TextReader tr = new StreamReader( local_pp.GetStream() ) )
      {
        local_custom_document = XDocument.Load( tr );
      }

      IEnumerable<int> pids =
      (
          from d in remote_custom_document.Root.Descendants()
          where d.Name.LocalName == "property"
          select int.Parse( d.Attribute( XName.Get( "pid" ) ).Value )
      );

      int pid = pids.Max() + 1;

      foreach( XElement remote_property in remote_custom_document.Root.Elements() )
      {
        bool found = false;
        foreach( XElement local_property in local_custom_document.Root.Elements() )
        {
          var remote_property_name = remote_property.Attribute( XName.Get( "name" ) );
          var local_property_name = local_property.Attribute( XName.Get( "name" ) );

          if( remote_property != null && local_property_name != null && remote_property_name.Value.Equals( local_property_name.Value ) )
          {
            found = true;
          }
        }

        if( !found )
        {
          remote_property.SetAttributeValue( XName.Get( "pid" ), pid );
          local_custom_document.Root.Add( remote_property );

          pid++;
        }
      }

      // Save the modified local custom styles.xml file.
      using( TextWriter tw = new StreamWriter( new PackagePartStream( local_pp.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        local_custom_document.Save( tw, SaveOptions.None );
      }
    }

    private void merge_numbering( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote )
    {
      // Add each remote numbering to this document.
      var abstractNumElement = _numbering.Root.Elements( XName.Get( "abstractNum", w.NamespaceName ) );
      var remote_abstractNums = remote._numbering.Root.Elements( XName.Get( "abstractNum", w.NamespaceName ) );

      int guidd = -1;
      foreach( var an in abstractNumElement )
      {
        var a = an.Attribute( XName.Get( "abstractNumId", w.NamespaceName ) );
        if( a != null )
        {
          int i;
          if( int.TryParse( a.Value, out i ) )
          {
            if( i > guidd )
            {
              guidd = i;
            }
          }
        }
      }
      guidd++;

      var remote_nums = remote._numbering.Root.Elements( XName.Get( "num", w.NamespaceName ) );
      var numElement = _numbering.Root.Elements( XName.Get( "num", w.NamespaceName ) );

      int guidd2 = 0;
      foreach( var an in numElement )
      {
        var a = an.Attribute( XName.Get( "numId", w.NamespaceName ) );
        if( a != null )
        {
          int i;
          if( int.TryParse( a.Value, out i ) )
          {
            if( i > guidd2 )
            {
              guidd2 = i;
            }
          }
        }
      }
      guidd2++;

      foreach( XElement remote_abstractNum in remote_abstractNums )
      {
        var currentGuidd2 = guidd2;
        var abstractNumId = remote_abstractNum.Attribute( XName.Get( "abstractNumId", w.NamespaceName ) );
        if( abstractNumId != null )
        {
          var abstractNumIdValue = abstractNumId.Value;
          abstractNumId.SetValue( guidd );

          foreach( XElement remote_num in remote_nums )
          {
            // in document
            var numIds = remote_mainDoc.Descendants( XName.Get( "numId", w.NamespaceName ) );
            foreach( var numId in numIds )
            {
              var attr = numId.Attribute( XName.Get( "val", w.NamespaceName ) );
              if( attr != null && attr.Value.Equals( remote_num.Attribute( XName.Get( "numId", w.NamespaceName ) ).Value ) )
              {
                attr.SetValue( currentGuidd2 );
              }
            }
            remote_num.SetAttributeValue( XName.Get( "numId", w.NamespaceName ), currentGuidd2 );
            //abstractNumId of this remote_num
            var e = remote_num.Element( XName.Get( "abstractNumId", w.NamespaceName ) );
            var a2 = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( a2 != null && a2.Value.Equals( abstractNumIdValue ) )
              a2.SetValue( guidd );
            currentGuidd2++;
          }
        }

        guidd++;
      }

      if( abstractNumElement != null )
      {
        if( abstractNumElement.Count() > 0 )
        {
          abstractNumElement.Last().AddAfterSelf( remote_abstractNums );
        }
        else
        {
          _numbering.Root.Add( remote_abstractNums );
        }
      }

      if( numElement != null )
      {
        if( numElement.Count() > 0 )
        {
          numElement.Last().AddAfterSelf( remote_nums );
        }
        else
        {
          _numbering.Root.Add( remote_nums );
        }
      }
    }

    private void merge_fonts( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote )
    {
      // Add each remote font to this document.
      IEnumerable<XElement> remote_fonts = remote._fontTable.Root.Elements( XName.Get( "font", DocX.w.NamespaceName ) );
      IEnumerable<XElement> local_fonts = _fontTable.Root.Elements( XName.Get( "font", DocX.w.NamespaceName ) );

      foreach( XElement remote_font in remote_fonts )
      {
        bool flag_addFont = true;
        foreach( XElement local_font in local_fonts )
        {
          if( local_font.Attribute( XName.Get( "name", DocX.w.NamespaceName ) ).Value == remote_font.Attribute( XName.Get( "name", DocX.w.NamespaceName ) ).Value )
          {
            flag_addFont = false;
            break;
          }
        }

        if( flag_addFont )
        {
          _fontTable.Root.Add( remote_font );
        }
      }
    }

    private void merge_styles( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_footnotes, XDocument remote_endnotes )
    {
      var local_styles = new Dictionary<string, string>();
      foreach( XElement local_style in _styles.Root.Elements( XName.Get( "style", w.NamespaceName ) ) )
      {
        var temp = new XElement( local_style );
        var styleId = temp.Attribute( XName.Get( "styleId", w.NamespaceName ) );
        var value = styleId.Value;
        styleId.Remove();
        var key = Regex.Replace( temp.ToString(), @"\s+", "" );
        if( !local_styles.ContainsKey( key ) )
        {
          local_styles.Add( key, value );
        }
      }

      // Add each remote style to this document.
      var remote_styles = remote._styles.Root.Elements( XName.Get( "style", w.NamespaceName ) );
      foreach( XElement remote_style in remote_styles )
      {
        var temp = new XElement( remote_style );
        var styleId = temp.Attribute( XName.Get( "styleId", w.NamespaceName ) );
        var value = styleId.Value;
        styleId.Remove();
        var key = Regex.Replace( temp.ToString(), @"\s+", "" );
        String guuid;

        // Check to see if the local document already contains the remote style.
        if( local_styles.ContainsKey( key ) )
        {
          String local_value;
          local_styles.TryGetValue( key, out local_value );

          // If the styleIds are the same then nothing needs to be done.
          if( local_value == value )
            continue;
          // All we need to do is update the styleId.
          guuid = local_value;
        }
        else
        {
          guuid = Guid.NewGuid().ToString();
          // Set the styleId in the remote_style to this new Guid
          remote_style.SetAttributeValue( XName.Get( "styleId", w.NamespaceName ), guuid );
        }

        foreach( XElement e in remote_mainDoc.Root.Descendants( XName.Get( "pStyle", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleId.Value ) )
          {
            e_styleId.SetValue( guuid );
          }
        }

        foreach( XElement e in remote_mainDoc.Root.Descendants( XName.Get( "rStyle", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleId.Value ) )
          {
            e_styleId.SetValue( guuid );
          }
        }

        foreach( XElement e in remote_mainDoc.Root.Descendants( XName.Get( "tblStyle", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleId.Value ) )
          {
            e_styleId.SetValue( guuid );
          }
        }

        if( remote_endnotes != null )
        {
          foreach( XElement e in remote_endnotes.Root.Descendants( XName.Get( "rStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleId.Value ) )
            {
              e_styleId.SetValue( guuid );
            }
          }

          foreach( XElement e in remote_endnotes.Root.Descendants( XName.Get( "pStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleId.Value ) )
            {
              e_styleId.SetValue( guuid );
            }
          }
        }

        if( remote_footnotes != null )
        {
          foreach( XElement e in remote_footnotes.Root.Descendants( XName.Get( "rStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleId.Value ) )
            {
              e_styleId.SetValue( guuid );
            }
          }

          foreach( XElement e in remote_footnotes.Root.Descendants( XName.Get( "pStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleId.Value ) )
            {
              e_styleId.SetValue( guuid );
            }
          }
        }

        // Make sure they don't clash by using a uuid.
        styleId.SetValue( guuid );
        _styles.Root.Add( remote_style );
      }
    }

    protected void clonePackageRelationship( DocX remote_document, PackagePart pp, XDocument remote_mainDoc )
    {
      var url = pp.Uri.OriginalString.Replace( "/", "" );
      var remote_rels = remote_document.PackagePart.GetRelationships();
      foreach( var remote_rel in remote_rels )
      {
        if( url.Equals( "word" + remote_rel.TargetUri.OriginalString.Replace( "/", "" ) ) )
        {
          var remote_Id = remote_rel.Id;
          var local_Id = this.PackagePart.CreateRelationship( remote_rel.TargetUri, remote_rel.TargetMode, remote_rel.RelationshipType ).Id;

          // Replace all instances of remote_Id in the local document with local_Id
          this.ReplaceAllRemoteID( remote_mainDoc, "blip", "embed", a.NamespaceName, remote_Id, local_Id );
          // Replace all instances of remote_Id in the local document with local_Id (for shapes)
          this.ReplaceAllRemoteID( remote_mainDoc, "imagedata", "id", v.NamespaceName, remote_Id, local_Id );
          break;
        }
      }
    }

    protected PackagePart clonePackagePart( PackagePart pp )
    {
      var new_pp = _package.CreatePart( pp.Uri, pp.ContentType, CompressionOption.Normal );

      using( Stream s_read = pp.GetStream() )
      {
        using( Stream s_write = new PackagePartStream( new_pp.GetStream( FileMode.Create ) ) )
        {
          CopyStream( s_read, s_write );
        }
      }

      return new_pp;
    }

    protected string GetMD5HashFromStream( Stream stream )
    {
      MD5 md5 = new MD5CryptoServiceProvider();
      byte[] retVal = md5.ComputeHash( stream );

      StringBuilder sb = new StringBuilder();
      for( int i = 0; i < retVal.Length; i++ )
      {
        sb.Append( retVal[ i ].ToString( "x2" ) );
      }
      return sb.ToString();
    }

    private static void PopulateDocument( DocX document, Package package )
    {
      var headers = new Headers();
      headers.Odd = document.GetHeaderByType( "default" );
      headers.Even = document.GetHeaderByType( "even" );
      headers.First = document.GetHeaderByType( "first" );

      var footers = new Footers();
      footers.Odd = document.GetFooterByType( "default" );
      footers.Even = document.GetFooterByType( "even" );
      footers.First = document.GetFooterByType( "first" );

      //// Get the sectPr for this document.
      //XElement sectPr = document.mainDoc.Descendants(XName.Get("sectPr", DocX.w.NamespaceName)).Single();

      //if (sectPr != null)
      //{
      //    // Extract the even header reference
      //    var header_even_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "headerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "even");
      //    string id = header_even_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
      //    var res = document.mainPart.GetRelationship(id);
      //    string ans = res.SourceUri.OriginalString;
      //    headers.even.xml_filename = ans;

      //    // Extract the odd header reference
      //    var header_odd_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "headerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "default");
      //    string id2 = header_odd_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
      //    var res2 = document.mainPart.GetRelationship(id2);
      //    string ans2 = res2.SourceUri.OriginalString;
      //    headers.odd.xml_filename = ans2;

      //    // Extract the first header reference
      //    var header_first_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "h
      //eaderReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "first");
      //    string id3 = header_first_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
      //    var res3 = document.mainPart.GetRelationship(id3);
      //    string ans3 = res3.SourceUri.OriginalString;
      //    headers.first.xml_filename = ans3;

      //    // Extract the even footer reference
      //    var footer_even_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "even");
      //    string id4 = footer_even_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
      //    var res4 = document.mainPart.GetRelationship(id4);
      //    string ans4 = res4.SourceUri.OriginalString;
      //    footers.even.xml_filename = ans4;

      //    // Extract the odd footer reference
      //    var footer_odd_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "default");
      //    string id5 = footer_odd_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
      //    var res5 = document.mainPart.GetRelationship(id5);
      //    string ans5 = res5.SourceUri.OriginalString;
      //    footers.odd.xml_filename = ans5;

      //    // Extract the first footer reference
      //    var footer_first_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "first");
      //    string id6 = footer_first_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
      //    var res6 = document.mainPart.GetRelationship(id6);
      //    string ans6 = res6.SourceUri.OriginalString;
      //    footers.first.xml_filename = ans6;

      //}

      document.Xml = document._mainDoc.Root.Element( w + "body" );
      document._headers = headers;
      document._footers = footers;
      document._settingsPart = HelperFunctions.CreateOrGetSettingsPart( package );

      var ps = package.GetParts();

      //document.endnotesPart = HelperFunctions.GetPart();
      foreach( var rel in document.PackagePart.GetRelationships() )
      {
        var uriString = "/word/" + rel.TargetUri.OriginalString.Replace( "/word/", "" ).Replace( "file://", "" );

        switch( rel.RelationshipType )
        {
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes":
            document._endnotesPart = package.GetPart( new Uri( uriString, UriKind.Relative ) );
            using( TextReader tr = new StreamReader( document._endnotesPart.GetStream() ) )
              document._endnotes = XDocument.Load( tr );
            break;

          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes":
            document._footnotesPart = package.GetPart( new Uri( uriString, UriKind.Relative ) );
            using( TextReader tr = new StreamReader( document._footnotesPart.GetStream() ) )
              document._footnotes = XDocument.Load( tr );
            break;

          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
            document._stylesPart = package.GetPart( new Uri( uriString, UriKind.Relative ) );
            using( TextReader tr = new StreamReader( document._stylesPart.GetStream() ) )
              document._styles = XDocument.Load( tr );
            break;

          case "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects":
            document._stylesWithEffectsPart = package.GetPart( new Uri( uriString, UriKind.Relative ) );
            using( TextReader tr = new StreamReader( document._stylesWithEffectsPart.GetStream() ) )
              document._stylesWithEffects = XDocument.Load( tr );
            break;

          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
            document._fontTablePart = package.GetPart( new Uri( uriString, UriKind.Relative ) );
            using( TextReader tr = new StreamReader( document._fontTablePart.GetStream() ) )
              document._fontTable = XDocument.Load( tr );
            break;

          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering":
            document._numberingPart = package.GetPart( new Uri( uriString, UriKind.Relative ) );
            using( TextReader tr = new StreamReader( document._numberingPart.GetStream() ) )
              document._numbering = XDocument.Load( tr );
            break;

          default:
            break;
        }
      }
    }

    private static void CopyStream( Stream input, Stream output, int bufferSize = 32768 )
    {
      byte[] buffer = new byte[ bufferSize ];
      int read;
      while( ( read = input.Read( buffer, 0, buffer.Length ) ) > 0 )
      {
        output.Write( buffer, 0, read );
      }
    }

    private string GetNextFreeRelationshipID()
    {
      int id =
        (
          from r in this.PackagePart.GetRelationships()
          where r.Id.Substring( 0, 3 ).Equals( "rId" )
          select int.Parse( r.Id.Substring( 3 ) )
        ).DefaultIfEmpty().Max();

      // The convension for ids is rid01, rid02, etc
      var newId = id.ToString();
      int result;
      if( int.TryParse( newId, out result ) )
        return ( "rId" + ( result + 1 ) );

      var guid = String.Empty;
      do
      {
        guid = Guid.NewGuid().ToString();
      } while( Char.IsDigit( guid[ 0 ] ) );
      return guid;
    }

    private byte[] MergeArrays( byte[] array1, byte[] array2 )
    {
      byte[] result = new byte[ array1.Length + array2.Length ];
      Buffer.BlockCopy( array2, 0, result, 0, array2.Length );
      Buffer.BlockCopy( array1, 0, result, array2.Length, array1.Length );
      return result;
    }

    private Hyperlink AddHyperlink( string text, Uri uri, string anchor )
    {
      var i = new XElement
      (
          XName.Get( "hyperlink", w.NamespaceName ),
          new XAttribute( r + "id", string.Empty ),
          new XAttribute( w + "history", "1" ),
          !string.IsNullOrEmpty( anchor ) ? new XAttribute( w + "anchor", anchor ) : null,
          new XElement( XName.Get( "r", w.NamespaceName ),
          new XElement( XName.Get( "rPr", w.NamespaceName ),
          new XElement( XName.Get( "rStyle", w.NamespaceName ),
          new XAttribute( w + "val", "Hyperlink" ) ) ),
          new XElement( XName.Get( "t", w.NamespaceName ), text ) )
      );

      var h = new Hyperlink( this, this.PackagePart, i );

      h.text = text;
      if( uri != null )
      {
        h.uri = uri;
      }

      this.AddHyperlinkStyleIfNotPresent();

      return h;
    }

    #endregion

    #region Constructors

    internal DocX( DocX document, XElement xml )
        : base( document, xml )
    {

    }

    #endregion

    #region IDisposable Members

    /// <summary>
    /// Releases all resources used by this document.
    /// </summary>
    /// <example>
    /// If you take advantage of the using keyword, Dispose() is automatically called for you.
    /// <code>
    /// // Load document.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///      // The document is only in memory while in this scope.
    ///
    /// }// Dispose() is automatically called at this point.
    /// </code>
    /// </example>
    /// <example>
    /// This example is equilivant to the one above example.
    /// <code>
    /// // Load document.
    /// DocX document = DocX.Load(@"C:\Example\Test.docx");
    /// 
    /// // Do something with the document here.
    ///
    /// // Dispose of the document.
    /// document.Dispose();
    /// </code>
    /// </example>
    public void Dispose()
    {
      _package.Close();
    }

    #endregion
  }
}
