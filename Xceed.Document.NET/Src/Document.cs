﻿/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


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
using System.Collections.ObjectModel;
using System.Net;
using System.Diagnostics;
using System.Reflection;
using System.ComponentModel;

namespace Xceed.Document.NET
{

  // When calling doc1.InsertDocument( doc2...), sets how to merge the documents styles and headers/footers.
  public enum MergingMode
  {
    Local,
    Remote,
    Both
  }

  public class Document : Container, IDisposable
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
    static internal XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    internal static XNamespace n = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    static internal XNamespace v = "urn:schemas-microsoft-com:vml";
    static internal XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    static internal XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    static internal XNamespace w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    static internal XNamespace w15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    static internal XNamespace o = "urn:schemas-microsoft-com:office:office";
    static internal XNamespace d = "http://www.w3.org/2000/09/xmldsig#";
    static internal XNamespace doffice = "http://schemas.microsoft.com/office/2006/digsig";
    static internal XNamespace w10 = "urn:schemas-microsoft-com:office:word";
    #endregion

    #region Private Members    

    private readonly object nextFreeDocPrIdLock = new object();
    private long? nextFreeDocPrId;
    private string _defaultParagraphStyleId;
    private static HashAlgorithmName _hashAlgorithmName;

    private static double DefaultFontSize = 11d;
    private static string DefaultFontFamily = "Calibri";
    private static string ModifiedByCoreProperty = "cp:lastModifiedBy";

    private int _lastEditId = -1;

    private IList<Section> _cachedSections;

    private static List<string> _imageContentTypes = new List<string>
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
                                                        "image/x-emf",
                                                        "image/wmf"
                                                    };

    #endregion

    #region Internal Constants

    internal const string RelationshipImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    internal const string RelationshipChart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
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
    internal bool _isCopyingDocument = false;

    #region Internal variables defined foreach Document object
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

    public override IList<Section> Sections
    {
      get
      {
        return _cachedSections;
      }
    }

    public NoteProperties EndnoteProperties
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].EndnoteProperties." );
        }
        return this.Sections[ 0 ].EndnoteProperties;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].EndnoteProperties." );
        }
        this.Sections[ 0 ].EndnoteProperties = value;
      }
    }

    public NoteProperties FootnoteProperties
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].FootnoteProperties." );
        }
        return this.Sections[ 0 ].FootnoteProperties;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].FootnoteProperties." );
        }
        this.Sections[ 0 ].FootnoteProperties = value;
      }
    }

    public float MarginTop
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginTop." );
        }
        return this.Sections[ 0 ].MarginTop;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginTop." );
        }
        this.Sections[ 0 ].MarginTop = value;
      }
    }

    public float MarginBottom
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginBottom." );
        }
        return this.Sections[ 0 ].MarginBottom;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginBottom." );
        }
        this.Sections[ 0 ].MarginBottom = value;
      }
    }

    public float MarginLeft
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginLeft." );
        }
        return this.Sections[ 0 ].MarginLeft;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginLeft." );
        }
        this.Sections[ 0 ].MarginLeft = value;
      }
    }

    public float MarginRight
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginRight." );
        }
        return this.Sections[ 0 ].MarginRight;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginRight." );
        }
        this.Sections[ 0 ].MarginRight = value;
      }
    }

    public float MarginHeader
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginHeader." );
        }
        return this.Sections[ 0 ].MarginHeader;
      }
      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginHeader." );
        }
        this.Sections[ 0 ].MarginHeader = value;
      }
    }

    public float MarginFooter
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginFooter." );
        }
        return this.Sections[ 0 ].MarginFooter;
      }
      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MarginFooter." );
        }
        this.Sections[ 0 ].MarginFooter = value;
      }
    }

    public bool MirrorMargins
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MirrorMargins." );
        }
        return this.Sections[ 0 ].MirrorMargins;
      }
      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].MirrorMargins." );
        }
        this.Sections[ 0 ].MirrorMargins = value;
      }
    }

    public float PageWidth
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageWidth." );
        }
        return this.Sections[ 0 ].PageWidth;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageWidth." );
        }
        this.Sections[ 0 ].PageWidth = value;
      }
    }

    public float PageHeight
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageHeight." );
        }
        return this.Sections[ 0 ].PageHeight;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageHeight." );
        }
        this.Sections[ 0 ].PageHeight = value;
      }
    }

    public Xceed.Drawing.Color PageBackground
    {
      get
      {
        var background = _mainDoc.Root.Element( XName.Get( "background", w.NamespaceName ) );
        if( background != null )
        {
          var color = background.Attribute( XName.Get( "color", Document.w.NamespaceName ) );
          if( color != null )
          {
            return HelperFunctions.GetColorFromHtml( color.Value );
          }
        }

        return Xceed.Drawing.Color.White;
      }

      set
      {
        var background = _mainDoc.Root.Element( XName.Get( "background", w.NamespaceName ) );
        if( background != null )
        {
          background.Remove();
        }
        background = new XElement( XName.Get( "background", Document.w.NamespaceName ) );
        _mainDoc.Root.AddFirst( background );

        background.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), value.ToHex() );
      }
    }

    public Borders PageBorders
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageBorders." );
        }
        return this.Sections[ 0 ].PageBorders;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageBorders." );
        }
        this.Sections[ 0 ].PageBorders = value;
      }
    }

    //public DocumentElement PageWatermark
    //{
    //  get
    //  {
    //    if( this.Sections.Count > 1 )
    //    {
    //      Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageWatermark." );
    //    }
    //    return this.Sections[ 0 ].PageWatermark;
    //  }

    //  set
    //  {
    //    if( this.Sections.Count > 1 )
    //    {
    //      Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageWatermark." );
    //    }
    //    this.Sections[ 0 ].PageWatermark = value;
    //  }
    //}

    public bool IsProtected
    {
      get
      {
        return _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).Count() > 0;
      }
    }

    public bool IsPasswordProtected
    {
      get
      {
        if( !IsProtected )
        {
          return false;
        }
        else
        {
          var documentProtection = _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) );
          return documentProtection.FirstOrDefault( x => x.Attribute( XName.Get( "cryptAlgorithmClass", w.NamespaceName ) ) != null ) != null &&
                 documentProtection.FirstOrDefault( x => x.Attribute( XName.Get( "cryptAlgorithmType", w.NamespaceName ) ) != null ) != null &&
                 documentProtection.FirstOrDefault( x => x.Attribute( XName.Get( "cryptAlgorithmSid", w.NamespaceName ) ) != null ) != null &&
                 documentProtection.FirstOrDefault( x => x.Attribute( XName.Get( "cryptSpinCount", w.NamespaceName ) ) != null ) != null &&
                 documentProtection.FirstOrDefault( x => x.Attribute( XName.Get( "cryptProviderType", w.NamespaceName ) ) != null ) != null &&
                 documentProtection.FirstOrDefault( x => x.Attribute( XName.Get( "hash", w.NamespaceName ) ) != null ) != null &&
                 documentProtection.FirstOrDefault( x => x.Attribute( XName.Get( "salt", w.NamespaceName ) ) != null ) != null;
        }
      }
    }

    public PageLayout PageLayout
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].PageLayout." );
        }
        return this.Sections[ 0 ].PageLayout;
      }
    }

    public Headers Headers
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].Headers." );
        }
        return this.Sections[ 0 ].Headers;
      }
    }

    public Footers Footers
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].Footers." );
        }
        return this.Sections[ 0 ].Footers;
      }
    }

    public bool DifferentOddAndEvenPages
    {
      get
      {
        var evenAndOddHeaders = _settings.Root.Element( w + "evenAndOddHeaders" );

        return ( evenAndOddHeaders != null );
      }

      set
      {
        var evenAndOddHeaders = _settings.Root.Element( w + "evenAndOddHeaders" );
        if( evenAndOddHeaders == null )
        {
          if( value )
          {
            _settings.Root.AddFirst( new XElement( w + "evenAndOddHeaders" ) );
          }
        }
        else
        {
          if( !value )
          {
            evenAndOddHeaders.Remove();
          }
        }
      }
    }

    public bool DifferentFirstPage
    {
      get
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].DifferentFirstPage." );
        }
        return this.Sections[ 0 ].DifferentFirstPage;
      }

      set
      {
        if( this.Sections.Count > 1 )
        {
          Debug.WriteLine( "This document contains more than 1 section. Consider using Sections[wantedSection].DifferentFirstPage." );
        }
        this.Sections[ 0 ].DifferentFirstPage = value;
      }
    }

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

    public IEnumerable<string> FootnotesText
    {
      get
      {
        foreach( var note in _footnotes.Root.Elements( w + "footnote" ) )
          yield return HelperFunctions.GetText( note );
      }
    }

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
        // In Body.
        // Faster to search the document.Xml instead of document.Paragraphs.
        var documentBookmarkStartsXml = this.Xml.Descendants( XName.Get( "bookmarkStart", Document.w.NamespaceName ) );
        foreach( var bookmarkStartXml in documentBookmarkStartsXml )
        {
          var bookmarkName = bookmarkStartXml.Attribute( XName.Get( "name", Document.w.NamespaceName ) ).Value;

          var paragraphXml = bookmarkStartXml.Parent;
          while( ( paragraphXml != null ) && ( paragraphXml.Name != XName.Get( "p", Document.w.NamespaceName ) ) )
          {
            paragraphXml = paragraphXml.Parent;
          }

          // bookmarkStart is not in a paragraph, get bookmarkEnd.
          if( paragraphXml == null )
          {
            var bookmarkId = bookmarkStartXml.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value;
            var bookmarkEnd = this.Xml.Descendants( XName.Get( "bookmarkEnd", Document.w.NamespaceName ) )
                                  .Where( x => x.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value == bookmarkId )
                                  .FirstOrDefault();

            if( bookmarkEnd == null )
              throw new InvalidDataException( "bookmarkEnd can't be found to match bookmarkStart with name " + bookmarkName + "." );

            paragraphXml = bookmarkEnd.Parent;
            while( ( paragraphXml != null ) && ( paragraphXml.Name != XName.Get( "p", Document.w.NamespaceName ) ) )
            {
              paragraphXml = paragraphXml.Parent;
            }

            // bookmarkStart is not in a paragraph, get bookmarkEnd.
            if( paragraphXml == null )
              throw new InvalidDataException( "bookmarkStart and bookmarkEnd are not part of a paragraph. This is currently not supported." );
          }

          bookmarks.Add( new Bookmark
          {
            Name = bookmarkName,
            Id = bookmarkStartXml.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value,
            Paragraph = new Paragraph( this, paragraphXml, -1 ) { PackagePart = this.PackagePart }
          } );
        }

        foreach( var section in this.Sections )
        {
          // In Headers.
          var headers = section.Headers;
          if( headers != null )
          {
            if( headers.Odd != null )
            {
              foreach( var paragraph in headers.Odd.Paragraphs )
              {
                bookmarks.AddRange( paragraph.GetBookmarks() );
              }
            }
            if( headers.Even != null )
            {
              foreach( var paragraph in headers.Even.Paragraphs )
              {
                bookmarks.AddRange( paragraph.GetBookmarks() );
              }
            }
            if( headers.First != null )
            {
              foreach( var paragraph in headers.First.Paragraphs )
              {
                bookmarks.AddRange( paragraph.GetBookmarks() );
              }
            }
          }

          // In Footers.
          var footers = section.Footers;
          if( footers != null )
          {
            if( footers.Odd != null )
            {
              foreach( var paragraph in footers.Odd.Paragraphs )
              {
                bookmarks.AddRange( paragraph.GetBookmarks() );
              }
            }

            if( footers.Even != null )
            {
              foreach( var paragraph in footers.Even.Paragraphs )
              {
                bookmarks.AddRange( paragraph.GetBookmarks() );
              }
            }

            if( footers.First != null )
            {
              foreach( var paragraph in footers.First.Paragraphs )
              {
                bookmarks.AddRange( paragraph.GetBookmarks() );
              }
            }
          }
        }

        return bookmarks;
      }
    }

    public List<Formatting> ParagraphFormattings
    {
      get
      {
        var paragraphStyles =
       (
           from s in _styles.Element( Document.w + "styles" ).Elements( Document.w + "style" )
           let type = s.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
           where ( ( type != null )
                    && ( ( type.Value == "paragraph" ) || ( type.Value == "character" ) ) )
           select s
       );

        var formattings = new List<Formatting>();

        foreach( var style in paragraphStyles )
        {
          var formatting = HelperFunctions.GetFormattingFromStyle( this, style.GetAttribute( XName.Get( "styleId", Document.w.NamespaceName ) ) );
          formattings.Add( formatting );
        }

        return formattings;
      }
    }























    #endregion

    #region Public Methods



    public override Section InsertSection( bool trackChanges )
    {
      return this.InsertSection( trackChanges, false );
    }

    public override Section InsertSectionPageBreak( bool trackChanges = false )
    {
      return this.InsertSection( trackChanges, true );
    }

    public override List<string> FindUniqueByPattern( string pattern, RegexOptions options )
    {
      // FindUniqueByPattern in the main body of document.
      var items = base.FindUniqueByPattern( pattern, options );

      // FindUniqueByPattern in Headers of the document.
      foreach( var section in this.Sections )
      {
        var headerList = new List<Header>() { section.Headers.First, section.Headers.Even, section.Headers.Odd };
        foreach( var h in headerList )
        {
          if( h != null )
          {
            items.AddRange( h.FindUniqueByPattern( pattern, options ) );
          }
        }
      }

      // FindUniqueByPattern in Footers of the document.
      foreach( var section in this.Sections )
      {
        var footerList = new List<Footer> { section.Footers.First, section.Footers.Even, section.Footers.Odd };
        foreach( var f in footerList )
        {
          if( f != null )
          {
            items.AddRange( f.FindUniqueByPattern( pattern, options ) );
          }
        }
      }

      return items;
    }

    [Obsolete( "ReplaceText() with many parameters is obsolete. Use ReplaceText() with a StringReplaceTextOptions parameter instead." )]
    public override void ReplaceText( string searchValue,
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
      var replaceTextOptions = new StringReplaceTextOptions()
      {
        SearchValue = searchValue,
        NewValue = newValue,
        TrackChanges = trackChanges,
        RegExOptions = options,
        NewFormatting = newFormatting,
        FormattingToMatch = matchFormatting,
        FormattingToMatchOptions = fo,
        EscapeRegEx = escapeRegEx,
        UseRegExSubstitutions = useRegExSubstitutions,
        RemoveEmptyParagraph = removeEmptyParagraph
      };

      // ReplaceText in the main body of document.
      base.ReplaceText( replaceTextOptions );

      // ReplaceText in Headers of the document.
      foreach( var section in this.Sections )
      {
        var headerList = new List<Header>() { section.Headers.First, section.Headers.Even, section.Headers.Odd };
        foreach( var h in headerList )
        {
          if( h != null )
          {
            foreach( var p in h.Paragraphs )
            {
              p.ReplaceText( replaceTextOptions );
            }
          }
        }
      }

      // ReplaceText in Footers of the document.
      foreach( var section in this.Sections )
      {
        var footerList = new List<Footer> { section.Footers.First, section.Footers.Even, section.Footers.Odd };
        foreach( var f in footerList )
        {
          if( f != null )
          {
            foreach( var p in f.Paragraphs )
            {
              p.ReplaceText( replaceTextOptions );
            }
          }
        }
      }
    }

    [Obsolete( "ReplaceText() with many parameters is obsolete. Use ReplaceText() with a FunctionReplaceTextOptions parameter instead." )]
    public override void ReplaceText( string searchValue,
                                      Func<string, string> regexMatchHandler,
                                      bool trackChanges = false,
                                      RegexOptions options = RegexOptions.None,
                                      Formatting newFormatting = null,
                                      Formatting matchFormatting = null,
                                      MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch,
                                      bool removeEmptyParagraph = true )
    {
      var replaceTextOptions = new FunctionReplaceTextOptions()
      {
        FindPattern = searchValue,
        RegexMatchHandler = regexMatchHandler,
        TrackChanges = trackChanges,
        RegExOptions = options,
        NewFormatting = newFormatting,
        FormattingToMatch = matchFormatting,
        FormattingToMatchOptions = fo,
        RemoveEmptyParagraph = removeEmptyParagraph
      };

      // Replace text in body of the Document.
      base.ReplaceText( replaceTextOptions );

      // Replace text in headers and footers of the Document.
      foreach( var section in this.Sections )
      {
        var headersFootersList = new List<IParagraphContainer>()
        {
          section.Headers.First,
          section.Headers.Even,
          section.Headers.Odd,
          section.Footers.First,
          section.Footers.Even,
          section.Footers.Odd,
        };

        foreach( var hf in headersFootersList )
        {
          if( hf != null )
          {
            foreach( var p in hf.Paragraphs )
            {
              p.ReplaceText( replaceTextOptions );
            }
          }
        }
      }
    }

    [Obsolete( "ReplaceText() with many parameters is obsolete. Use ReplaceText() with an ObjectReplaceTextOptions parameter instead." )]
    public override void ReplaceTextWithObject( string searchValue,
                                                DocumentElement objectToAdd,
                                                bool trackChanges = false,
                                                RegexOptions options = RegexOptions.None,
                                                Formatting matchFormatting = null,
                                                MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch,
                                                bool escapeRegEx = true,
                                                bool removeEmptyParagraph = true )
    {
      var replaceTextOptions = new ObjectReplaceTextOptions()
      {
        SearchValue = searchValue,
        NewObject = objectToAdd,
        TrackChanges = trackChanges,
        RegExOptions = options,
        FormattingToMatch = matchFormatting,
        FormattingToMatchOptions = fo,
        EscapeRegEx = escapeRegEx,
        RemoveEmptyParagraph = removeEmptyParagraph
      };

      base.ReplaceTextWithObject( replaceTextOptions );

      // ReplaceText in Headers of the document.
      foreach( var section in this.Sections )
      {
        var headerList = new List<Header>() { section.Headers.First, section.Headers.Even, section.Headers.Odd };
        foreach( var h in headerList )
        {
          if( h != null )
          {
            foreach( var p in h.Paragraphs )
            {
              p.ReplaceTextWithObject( replaceTextOptions );
            }
          }
        }
      }

      // ReplaceText in Footers of the document.
      foreach( var section in this.Sections )
      {
        var footerList = new List<Footer> { section.Footers.First, section.Footers.Even, section.Footers.Odd };
        foreach( var f in footerList )
        {
          if( f != null )
          {
            foreach( var p in f.Paragraphs )
            {
              p.ReplaceTextWithObject( replaceTextOptions );
            }
          }
        }
      }
    }

    public override bool ReplaceText( StringReplaceTextOptions replaceTextOptions )
    {
      bool replaceSuccess = false;

      // ReplaceText in the main body of document.
      if( ( replaceTextOptions.ContainerLocation & ReplaceTextContainer.Body ) == ReplaceTextContainer.Body )
      {
        replaceSuccess = base.ReplaceText( replaceTextOptions );
      }

      foreach( var section in this.Sections )
      {
        // ReplaceText in Headers of the document.
        if( ( replaceTextOptions.ContainerLocation & ReplaceTextContainer.Header ) == ReplaceTextContainer.Header )
        {
          var headerList = new List<Header>() { section.Headers.First, section.Headers.Even, section.Headers.Odd };
          foreach( var h in headerList )
          {
            if( h != null )
            {
              var result = h.ReplaceText( replaceTextOptions );

              if( !replaceSuccess )
              {
                replaceSuccess = result;
              }
            }
          }
        }

        // ReplaceText in Footers of the document.
        if( ( replaceTextOptions.ContainerLocation & ReplaceTextContainer.Footer ) == ReplaceTextContainer.Footer )
        {
          var footerList = new List<Footer> { section.Footers.First, section.Footers.Even, section.Footers.Odd };
          foreach( var f in footerList )
          {
            if( f != null )
            {
              var result = f.ReplaceText( replaceTextOptions );

              if( !replaceSuccess )
              {
                replaceSuccess = result;
              }
            }
          }
        }
      }

      return replaceSuccess;
    }

    public override bool ReplaceText( FunctionReplaceTextOptions replaceTextOptions )
    {
      bool replaceSuccess = false;

      // Replace text in body of the Document.
      if( ( replaceTextOptions.ContainerLocation & ReplaceTextContainer.Body ) == ReplaceTextContainer.Body )
      {
        replaceSuccess = base.ReplaceText( replaceTextOptions );
      }

      foreach( var section in this.Sections )
      {
        // ReplaceText in Headers of the document.
        if( ( replaceTextOptions.ContainerLocation & ReplaceTextContainer.Header ) == ReplaceTextContainer.Header )
        {
          var headerList = new List<Header>() { section.Headers.First, section.Headers.Even, section.Headers.Odd };
          foreach( var h in headerList )
          {
            if( h != null )
            {
              var result = h.ReplaceText( replaceTextOptions );

              if( !replaceSuccess )
              {
                replaceSuccess = result;
              }
            }
          }
        }

        // ReplaceText in Footers of the document.
        if( ( replaceTextOptions.ContainerLocation & ReplaceTextContainer.Footer ) == ReplaceTextContainer.Footer )
        {
          var footerList = new List<Footer> { section.Footers.First, section.Footers.Even, section.Footers.Odd };
          foreach( var f in footerList )
          {
            if( f != null )
            {
              var result = f.ReplaceText( replaceTextOptions );

              if( !replaceSuccess )
              {
                replaceSuccess = result;
              }
            }
          }
        }
      }

      return replaceSuccess;
    }

    public override bool ReplaceTextWithObject( ObjectReplaceTextOptions replaceTextOptions )
    {
      bool replaceSuccess = false;

      // Replace text in body of the Document.
      if( ( replaceTextOptions.ContainerLocation & ReplaceTextContainer.Body ) == ReplaceTextContainer.Body )
      {
        replaceSuccess = base.ReplaceTextWithObject( replaceTextOptions );
      }

      // ReplaceText in Headers of the document.
      foreach( var section in this.Sections )
      {
        var headerList = new List<Header>() { section.Headers.First, section.Headers.Even, section.Headers.Odd };
        foreach( var h in headerList )
        {
          if( h != null )
          {
            var result = h.ReplaceTextWithObject( replaceTextOptions );

            if( !replaceSuccess )
            {
              replaceSuccess = result;
            }
          }
        }

        // ReplaceText in Footers of the document.
        if( ( replaceTextOptions.ContainerLocation & ReplaceTextContainer.Footer ) == ReplaceTextContainer.Footer )
        {
          var footerList = new List<Footer> { section.Footers.First, section.Footers.Even, section.Footers.Odd };
          foreach( var f in footerList )
          {
            if( f != null )
            {
              var result = f.ReplaceTextWithObject( replaceTextOptions );

              if( !replaceSuccess )
              {
                replaceSuccess = result;
              }
            }
          }
        }
      }

      return replaceSuccess;
    }


















    public override void InsertAtBookmark( string toInsert, string bookmarkName, Formatting formatting = null )
    {
      // Insert in body of document.
      base.InsertAtBookmark( toInsert, bookmarkName, formatting );

      // Insert in headers/footers of document.
      foreach( var section in this.Sections )
      {
        var headerCollection = section.Headers;
        var headers = new List<Header> { headerCollection.First, headerCollection.Even, headerCollection.Odd };
        foreach( var header in headers.Where( x => x != null ) )
        {
          foreach( var paragraph in header.Paragraphs )
          {
            paragraph.InsertAtBookmark( toInsert, bookmarkName, formatting );
          }
        }

        var footerCollection = section.Footers;
        var footers = new List<Footer> { footerCollection.First, footerCollection.Even, footerCollection.Odd };
        foreach( var footer in footers.Where( x => x != null ) )
        {
          foreach( var paragraph in footer.Paragraphs )
          {
            paragraph.InsertAtBookmark( toInsert, bookmarkName, formatting );
          }
        }
      }
    }

    public override string[] ValidateBookmarks( params string[] bookmarkNames )
    {
      // Validate in body of document.
      var result = base.ValidateBookmarks( bookmarkNames ).ToList();

      foreach( var bookmarkName in bookmarkNames )
      {
        // Validate in headers/footers of document.
        foreach( var section in this.Sections )
        {
          var headers = new[] { section.Headers.First, section.Headers.Even, section.Headers.Odd }.Where( h => h != null ).ToList();
          var footers = new[] { section.Footers.First, section.Footers.Even, section.Footers.Odd }.Where( f => f != null ).ToList();

          if( headers.SelectMany( h => h.Paragraphs ).Any( p => p.ValidateBookmark( bookmarkName ) ) )
            return new string[ 0 ];
          if( footers.SelectMany( h => h.Paragraphs ).Any( p => p.ValidateBookmark( bookmarkName ) ) )
            return new string[ 0 ];
        }

        result.Add( bookmarkName );
      }

      return result.ToArray();
    }

    // Returns the name of the first occurence of a paragraph's style, with name == styleName.
    public static string GetParagraphStyleIdFromStyleName( Document document, string styleName )
    {
      if( string.IsNullOrEmpty( styleName ) || ( document == null ) )
        return null;

      // Load _styles if not loaded.
      if( document._styles == null )
      {
        var word_styles = document._package.GetPart( new Uri( "/word/styles.xml", UriKind.Relative ) );
        using( var tr = new StreamReader( word_styles.GetStream() ) )
        {
          document._styles = XDocument.Load( tr );
        }
      }

      // Check if this Paragraph StyleName exists in _styles.
      var paragraphStyle = HelperFunctions.GetParagraphStyleFromStyleName( document, styleName );
      if( paragraphStyle != null )
      {
        var styleId = paragraphStyle.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) );
        if( styleId != null )
          return styleId.Value;
      }

      return null;
    }

    public EditRestrictions GetProtectionType()
    {
      if( IsProtected )
      {
        XElement documentProtection = _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).FirstOrDefault();
        string edit_type = documentProtection.Attribute( XName.Get( "edit", w.NamespaceName ) ).Value;
        return (EditRestrictions)Enum.Parse( typeof( EditRestrictions ), edit_type );
      }

      return EditRestrictions.none;
    }

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

    public void RemoveProtection()
    {
      if( IsPasswordProtected )
        throw new UnauthorizedAccessException( "The document is password protected." );

      // Remove every node of type documentProtection.
      _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).Remove();
    }

    public void InsertDocument( Document remote_document, bool append = true, bool useSectionBreak = true, MergingMode mergingMode = MergingMode.Both )
    {
      XDocument remote_mainDoc;
      XElement remote_body, local_body;
      PackagePartCollection ppc;
      if( remote_document != null )
      {
        remote_document.SaveHeadersFooters();
      }
      MergeDocumentsPackages( remote_document, mergingMode, out remote_mainDoc, out remote_body, out local_body, out ppc );

      ManageSectionBreak( append, useSectionBreak, remote_body, local_body );

      MergeDocumentsRelations( remote_document, remote_mainDoc, remote_body, ppc );

      // Add the remote documents contents to this document.
      if( append )
      {
        local_body.Add( remote_body.Elements() );
      }
      else
      {
        local_body.AddFirst( remote_body.Elements() );
      }

      AddRootAttributes( remote_mainDoc );

      // Update the _cachedSection by reading the Xml to build new Sections.
      this.UpdateCacheSections();
    }







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

      var t = new Table( this, HelperFunctions.CreateTable( rowCount, columnCount, this.GetAvailableWidth() ), this.PackagePart );
      t.PackagePart = this.PackagePart;
      return t;
    }




    public new Table InsertTable( int index, Table t )
    {
      var t2 = base.InsertTable( index, t );
      t2.PackagePart = this.PackagePart;
      return t2;
    }

    public new Table InsertTable( Table t )
    {
      t = base.InsertTable( t );
      t.PackagePart = this.PackagePart;
      return t;
    }

    public new Table InsertTable( int index, int rowCount, int columnCount )
    {
      if( rowCount < 1 || columnCount < 1 )
        throw new ArgumentOutOfRangeException( "Row and Column count must be greater than zero." );

      var t = base.InsertTable( index, rowCount, columnCount );

      t.PackagePart = this.PackagePart;
      return t;
    }










    public Footnote AddFootnote( object footnoteContent, Formatting formatting = null )
    {
      if( footnoteContent == null )
        return null;

      var footnotes_package_uri = new Uri( "/word/footnotes.xml", UriKind.Relative );

      if( !this.Document._package.PartExists( footnotes_package_uri ) )
      {
        var mainDocPart = HelperFunctions.GetMainDocumentPart( this.Document._package );
        mainDocPart.CreateRelationship( footnotes_package_uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" );

        // Create default footnotes in _footnotes.
        this.Document._footnotes = XDocument.Parse
           ( @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                  <w:footnotes xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:footnote w:id=""-1"" w:type=""separator"">
                      <w:p>
                        <w:pPr>
                          <w:spacing w:lineRule=""auto"" w:line=""240"" w:after=""0""/>
                        </w:pPr>
                        <w:r>
                          <w:separator/>
                        </w:r>
                      </w:p>
                    </w:footnote>
                    <w:footnote w:id=""0"" w:type=""continuationSeparator"">
                      <w:p>
                        <w:pPr>
                          <w:spacing w:lineRule=""auto"" w:line=""240"" w:after=""0""/>
                        </w:pPr>
                        <w:r>
                          <w:continuationSeparator/>
                        </w:r>
                      </w:p>
                    </w:footnote>
                  </w:footnotes>"
           );

        // Create Footnotes package.
        var footnotes_package = this.Document._package.CreatePart( footnotes_package_uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", CompressionOption.Maximum );
        this.Document._footnotesPart = footnotes_package;
      }

      var newFootnoteId = this.GetMaxId( this.Document._footnotes.Root.LastNode as XElement );
      if( newFootnoteId < 0 )
        return null;

      // Create new footnote.
      var newFootnoteXml = XElement.Parse
         ( string.Format( @"<w:footnote w:id=""{0}"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">                             
                            </w:footnote>", newFootnoteId )
         );

      // Create new footnote paragraph.
      var newFootnoteParagraphXml = XElement.Parse(
                                    @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                                        <w:pPr>
                                          <w:pStyle w:val=""FootnoteText""/>
                                        </w:pPr>
                                        <w:r>
                                          <w:rPr>
                                            <w:rStyle w:val=""FootnoteReference""/>
                                          </w:rPr>
                                          <w:footnoteRef/>
                                        </w:r>
                                        <w:r>
                                          <w:t xml:space=""preserve""> </w:t>
                                        </w:r>
                                      </w:p>" );
      var newFootnoteParagraph = new Paragraph( this, newFootnoteParagraphXml, 0 ) { PackagePart = _footnotesPart };

      this.UpdateNoteXmlFromContent( newFootnoteXml, newFootnoteParagraph, footnoteContent, formatting );

      // Create new footnote from updated newFootnoteXml.
      var newFootnote = new Footnote( this, _footnotesPart, newFootnoteXml );

      // Add newFootnote in _footnotes.
      this.Document._footnotes.Root.Add( newFootnote.Xml );

      return newFootnote;
    }

    public Endnote AddEndnote( object endnoteContent, Formatting formatting = null )
    {
      if( endnoteContent == null )
        return null;

      var endnotes_package_uri = new Uri( "/word/endnotes.xml", UriKind.Relative );

      if( !this.Document._package.PartExists( endnotes_package_uri ) )
      {
        var mainDocPart = HelperFunctions.GetMainDocumentPart( this.Document._package );
        mainDocPart.CreateRelationship( endnotes_package_uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" );

        // Create default endnotes in _endnotes.
        this.Document._endnotes = XDocument.Parse
           ( @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                  <w:endnotes xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:endnote w:id=""-1"" w:type=""separator"">
                      <w:p>
                        <w:pPr>
                          <w:spacing w:lineRule=""auto"" w:line=""240"" w:after=""0""/>
                        </w:pPr>
                        <w:r>
                          <w:separator/>
                        </w:r>
                      </w:p>
                    </w:endnote>
                    <w:endnote w:id=""0"" w:type=""continuationSeparator"">
                      <w:p>
                        <w:pPr>
                          <w:spacing w:lineRule=""auto"" w:line=""240"" w:after=""0""/>
                        </w:pPr>
                        <w:r>
                          <w:continuationSeparator/>
                        </w:r>
                      </w:p>
                    </w:endnote>
                  </w:endnotes>"
           );

        // Create Endnotes package.
        var endnotes_package = this.Document._package.CreatePart( endnotes_package_uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml", CompressionOption.Maximum );
        this.Document._endnotesPart = endnotes_package;
      }

      var newEndnoteId = this.GetMaxId( this.Document._endnotes.Root.LastNode as XElement );
      if( newEndnoteId < 0 )
        return null;

      // Create new endnote.
      var newEndnoteXml = XElement.Parse
         ( string.Format( @"<w:endnote w:id=""{0}"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">                             
                            </w:endnote>", newEndnoteId )
         );

      // Create new endnote paragraph.
      var newEndnoteParagraphXml = XElement.Parse(
                                    @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                                        <w:pPr>
                                          <w:pStyle w:val=""EndnoteText""/>
                                        </w:pPr>
                                        <w:r>
                                          <w:rPr>
                                            <w:rStyle w:val=""EndnoteReference""/>
                                          </w:rPr>
                                          <w:endnoteRef/>
                                        </w:r>
                                        <w:r>
                                          <w:t xml:space=""preserve""> </w:t>
                                        </w:r>
                                      </w:p>" );
      var newEndnoteParagraph = new Paragraph( this, newEndnoteParagraphXml, 0 ) { PackagePart = _endnotesPart };

      this.UpdateNoteXmlFromContent( newEndnoteXml, newEndnoteParagraph, endnoteContent, formatting );

      // Create new endnote from updated newEndnoteXml.
      var newEndnote = new Endnote( this, _endnotesPart, newEndnoteXml );

      // Add newEndnote in _endnotes.
      this.Document._endnotes.Root.Add( newEndnote.Xml );

      return newEndnote;
    }

    public void ApplyTemplate( string templateFilePath )
    {
      ApplyTemplate( templateFilePath, true );
    }

    public void ApplyTemplate( string templateFilePath, bool includeContent )
    {
      if( !File.Exists( templateFilePath ) )
      {
        throw new FileNotFoundException( string.Format( "File could not be found {0}", templateFilePath ) );
      }
      using( FileStream packageStream = new FileStream( templateFilePath, FileMode.Open, FileAccess.Read ) )
      {
        if( packageStream.Length == 0 )
        {
          throw new FileNotFoundException( string.Format( "File should not be empty {0}", templateFilePath ) );
        }
        ApplyTemplate( packageStream, includeContent );
      }
    }

    public void ApplyTemplate( Stream templateStream )
    {
      ApplyTemplate( templateStream, true );
    }

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
              using( var stream = packagePart.GetStream( FileMode.Open, FileAccess.Read ) )
              {
                using( var xr = XmlReader.Create( stream ) )
                {
                  documentDoc = XDocument.Load( xr );
                }
              }
              break;
            case "/_rels/.rels":
              if( !_package.PartExists( packagePart.Uri ) )
              {
                _package.CreatePart( packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption );
              }
              var globalRelsPart = _package.GetPart( packagePart.Uri );
              using( var stream = packagePart.GetStream( FileMode.Open, FileAccess.Read ) )
              {
                using( var tr = new StreamReader( stream, Encoding.UTF8 ) )
                {
                  using( var globalRelsPartStream = new PackagePartStream( globalRelsPart.GetStream( FileMode.Create, FileAccess.Write ) ) )
                  {
                    using( var tw = new StreamWriter( globalRelsPartStream, Encoding.UTF8 ) )
                    {
                      tw.Write( tr.ReadToEnd() );
                    }
                  }
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

              var nativePart = _package.GetPart( packagePart.Uri );
              if( packagePart.Uri.ToString().EndsWith( ".xml" ) || packagePart.Uri.ToString().EndsWith( ".rels" ) )
              {
                var packagePartEncoding = Encoding.UTF8;

                using( var stream = packagePart.GetStream( FileMode.Open, FileAccess.Read ) )
                {
                  using( var tr = new StreamReader( stream, packagePartEncoding ) )
                  {
                    using( var nativePartStream = new PackagePartStream( nativePart.GetStream( FileMode.Create, FileAccess.Write ) ) )
                    {
                      using( var tw = new StreamWriter( nativePartStream, tr.CurrentEncoding ) )
                      {
                        tw.Write( tr.ReadToEnd() );
                      }
                    }
                  }
                }
              }
              else
              {
                using( var s_read = packagePart.GetStream() )
                {
                  using( var s_write = new PackagePartStream( nativePart.GetStream( FileMode.Create ) ) )
                  {
                    HelperFunctions.CopyStream( s_read, s_write );
                  }
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
          var documentNewPart = _package.CreatePart( documentPart.Uri, mainContentType, documentPart.CompressionOption );
          using( var stream = new PackagePartStream( documentNewPart.GetStream( FileMode.Create, FileAccess.Write ) ) )
          {
            using( XmlWriter xw = XmlWriter.Create( stream ) )
            {
              documentDoc.WriteTo( xw );
            }
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
          foreach( Paragraph paragraph in this.Paragraphs.ToList() )
          {
            paragraph.Remove( false );
          }
        }
      }
      finally
      {
        templatePackage.Close();
        PopulateDocument( Document, _package );
      }
    }

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

      return this.AddImage( filename as object, contentType );
    }

    public Image AddImage( Stream stream, string contentType = "image/jpeg" )
    {
      stream.Position = 0;
      return AddImage( stream as object, contentType );
    }




    public Hyperlink AddHyperlink( string text, Uri uri )
    {
      return this.AddHyperlinkCore( text, uri, null, null, null );
    }

    public Hyperlink AddHyperlink( string text, string anchor )
    {
      return this.AddHyperlinkCore( text, null, anchor, null, null );
    }

    public void AddHeaders()
    {
      this.Sections[ 0 ].AddHeadersOrFootersXml( true );
    }

    public void AddFooters()
    {
      this.Sections[ 0 ].AddHeadersOrFootersXml( false );
    }

    public virtual void Save( string password = "" )
    {
    }

    public virtual void SaveAs( string filename, string password = "" )
    {
      this.SetFileName( filename );
      _stream = null;
      Save( password );
    }

    public virtual void SaveAs( Stream stream, string password = "" )
    {
      if( !_isCopyingDocument )
      {
        this.SetFileName( null );
      }
      _stream = stream;
      Save( password );
    }

    public void AddCoreProperty( string propertyName, string propertyValue )
    {
      var propertyNamespacePrefix = propertyName.Contains( ":" ) ? propertyName.Split( ':' )[ 0 ] : "cp";
      var propertyLocalName = propertyName.Contains( ":" ) ? propertyName.Split( ':' )[ 1 ] : propertyName;

      // If this document does not contain a coreFilePropertyPart create one.)
      if( !_package.PartExists( new Uri( "/docProps/core.xml", UriKind.Relative ) ) )
        HelperFunctions.CreateCorePropertiesPart( this );

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
      Document.UpdateCorePropertyValue( this, propertyLocalName, propertyValue );
    }

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
      Document.UpdateCustomPropertyValue( this, cp );
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

      // Inserting a paragraph at a specific index needs an update of Paragraph cache.
      this.ClearParagraphsCache();

      return p;
    }

    public override Paragraph InsertParagraph( Paragraph p )
    {
      // copy paragraph's pictures and styles.
      // This will modify the p's Picture Ids before adding it in this Document with the following base.InsertParagraph( updatedParagraph );
      var updatedParagraph = this.InsertParagraphPictures( p );

      this.InsertParagraphStyles( updatedParagraph );      

      var newParagraph = base.InsertParagraph( updatedParagraph );
      newParagraph.PackagePart = this.PackagePart;

      return newParagraph;
    }

    public override Paragraph InsertParagraph( int index, Paragraph p )
    {
      // copy paragraph's pictures and styles.
      // This will modify the p's Picture Ids before adding it in this Document with the following base.InsertParagraph( index, updatedParagraph );
      var updatedParagraph = this.InsertParagraphPictures( p );

      this.InsertParagraphStyles( updatedParagraph );

      var returnParagraph = base.InsertParagraph( index, updatedParagraph );
      returnParagraph.PackagePart = this.PackagePart;

      // Inserting a paragraph at a specific index needs an update of Paragraph cache.
      this.ClearParagraphsCache();

      return returnParagraph;
    }

    public override Paragraph InsertParagraph( int index, string text, bool trackChanges, Formatting formatting )
    {
      var p = base.InsertParagraph( index, text, trackChanges, formatting );
      p.PackagePart = this.PackagePart;

      // Inserting a paragraph at a specific index needs an update of Paragraph cache.
      this.ClearParagraphsCache();

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

    public override Paragraph InsertEquation( String equation, Alignment align = Alignment.center )
    {
      var p = base.InsertEquation( equation, align );
      p.PackagePart = this.PackagePart;

      return p;
    }

    public void InsertChart( Chart chart, float width = 432f, float height = 252f )
    {
      this.InsertChart( chart, null, width, height );
    }

    public void InsertChartAfterParagraph( Chart chart, Paragraph paragraph, float width = 432f, float height = 252f )
    {
      this.InsertChart( chart, paragraph, width, height );
    }

    public List AddList( string listText = null, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false, Formatting formatting = null )
    {
      if( startNumber.HasValue && continueNumbering )
        throw new InvalidOperationException( "Cannot specify a start number and at the same time continue numbering from another list" );

      var list = new List( this, null );
      var result = HelperFunctions.CreateItemInList( list, listText, level, listType, startNumber, trackChanges, continueNumbering, formatting );
      var lastItem = result.Items.LastOrDefault();

      if( lastItem != null )
      {
        lastItem.PackagePart = this.PackagePart;
      }

      return result;
    }










    public List AddListItem( List list,
                             string listText,
                             int level = 0,
                             ListItemType listType = ListItemType.Numbered,
                             int? startNumber = null,
                             bool trackChanges = false,
                             bool continueNumbering = false,
                             Formatting formatting = null )
    {
      if( startNumber.HasValue && continueNumbering )
        throw new InvalidOperationException( "Cannot specify a start number and at the same time continue numbering from another list" );

      var result = HelperFunctions.CreateItemInList( list, listText, level, listType, startNumber, trackChanges, continueNumbering, formatting );
      var lastItem = result.Items.LastOrDefault();

      if( lastItem != null )
      {
        lastItem.PackagePart = this.PackagePart;
      }

      return result;
    }

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

    public new List InsertList( int index, List list )
    {
      base.InsertList( index, list );
      return list;
    }

    public TableOfContents InsertDefaultTableOfContents()
    {
      var switchesDictionary = TableOfContents.BuildTOCSwitchesDictionary( TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U );
      return InsertTableOfContents( "Table of contents", switchesDictionary );
    }

    public TableOfContents InsertTableOfContents( string title, IDictionary<TableOfContentsSwitches, string> switches, string headerStyle = null, int? rightTabPos = null )
    {
      var toc = TableOfContents.CreateTableOfContents( this, title, switches, headerStyle, rightTabPos );
      this.AddElementInXml( toc.Xml );
      return toc;
    }

    [Obsolete( "This method is obsolete and should no longer be used. Use the InsertTableOfContents methods containing an IDictionary<TableOfContentsSwitches, string> parameter instead." )]
    public TableOfContents InsertTableOfContents( string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null )
    {
      var switchesDictionary = TableOfContents.BuildTOCSwitchesDictionary( switches, maxIncludeLevel );

      var toc = TableOfContents.CreateTableOfContents( this, title, switchesDictionary, headerStyle, rightTabPos );
      this.AddElementInXml( toc.Xml );
      return toc;
    }

    [Obsolete( "This method is obsolete and should no longer be used. Use the InsertTableOfContents methods containing an IDictionary<TableOfContentsSwitches, string> parameter instead." )]
    public TableOfContents InsertTableOfContents( Paragraph reference, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null )
    {
      var switchesDictionary = TableOfContents.BuildTOCSwitchesDictionary( switches, maxIncludeLevel );
      var toc = TableOfContents.CreateTableOfContents( this, title, switchesDictionary, headerStyle, rightTabPos );
      reference.Xml.AddBeforeSelf( toc.Xml );
      return toc;
    }

    public TableOfContents InsertTableOfContents( Paragraph reference, string title, IDictionary<TableOfContentsSwitches, string> switches, string headerStyle = null, int? rightTabPos = null )
    {
      var toc = TableOfContents.CreateTableOfContents( this, title, switches, headerStyle, rightTabPos );
      reference.Xml.AddBeforeSelf( toc.Xml );
      return toc;
    }

    public virtual Document Copy()
    {
      return null;
    }

    public void AddPasswordProtection( EditRestrictions editRestrictions, string password, HashAlgorithm hashAlgorithm = HashAlgorithm.None )
    {
      if( IsPasswordProtected )
        throw new UnauthorizedAccessException( "Document is already password protected." );

      if( string.IsNullOrEmpty( password ) )
        return;

      // If no EditRestrictions, nothing to do
      if( editRestrictions == EditRestrictions.none )
        return;

      if( hashAlgorithm == HashAlgorithm.None || _hashAlgorithmName == HashAlgorithmName.NotSupported )
      {
        hashAlgorithm = HashAlgorithm.Sha512;
      }

      var cryptAlgorithmSid = "";
      var cryptProviderType = "";

      switch( hashAlgorithm )
      {
        case HashAlgorithm.Md2:
          cryptProviderType = "rsaFull";
          cryptAlgorithmSid = "1";
          _hashAlgorithmName = HashAlgorithmName.SHA1;
          break;
        case HashAlgorithm.Md5:
          cryptProviderType = "rsaFull";
          cryptAlgorithmSid = "3";
          _hashAlgorithmName = HashAlgorithmName.SHA1;
          break;
        case HashAlgorithm.Sha1:
          cryptProviderType = "rsaFull";
          cryptAlgorithmSid = "4";
          _hashAlgorithmName = HashAlgorithmName.SHA1Managed;
          break;
        case HashAlgorithm.Sha256:
          cryptProviderType = "rsaAES";
          cryptAlgorithmSid = "12";
          _hashAlgorithmName = HashAlgorithmName.SHA256Managed;
          break;
        case HashAlgorithm.Sha384:
          cryptProviderType = "rsaAES";
          cryptAlgorithmSid = "13";
          _hashAlgorithmName = HashAlgorithmName.SHA384Managed;
          break;
        case HashAlgorithm.Sha512:
          cryptProviderType = "rsaAES";
          cryptAlgorithmSid = "14";
          _hashAlgorithmName = HashAlgorithmName.SHA512Managed;
          break;
        default:
          throw new NotSupportedException( string.Format( "Hash algorithm '{0}' is not supported for document write protection.", hashAlgorithm ) );
      }

      // Init DocumentProtection element
      var documentProtection = new XElement( XName.Get( "documentProtection", w.NamespaceName ) );
      documentProtection.Add( new XAttribute( XName.Get( "edit", w.NamespaceName ), editRestrictions.ToString() ) );
      documentProtection.Add( new XAttribute( XName.Get( "enforcement", w.NamespaceName ), "1" ) );

      Encryption encryption = new Encryption();
      var encryptionObj = encryption.Encrypt( password, _hashAlgorithmName );

      documentProtection.Add( new XAttribute( XName.Get( "cryptProviderType", w.NamespaceName ), cryptProviderType ) );
      documentProtection.Add( new XAttribute( XName.Get( "cryptAlgorithmClass", w.NamespaceName ), "hash" ) );
      documentProtection.Add( new XAttribute( XName.Get( "cryptAlgorithmType", w.NamespaceName ), "typeAny" ) );
      documentProtection.Add( new XAttribute( XName.Get( "cryptAlgorithmSid", w.NamespaceName ), cryptAlgorithmSid ) );
      documentProtection.Add( new XAttribute( XName.Get( "cryptSpinCount", w.NamespaceName ), encryptionObj.Iterations.ToString() ) );
      documentProtection.Add( new XAttribute( XName.Get( "hash", w.NamespaceName ), Convert.ToBase64String( encryptionObj.KeyValues ) ) );
      documentProtection.Add( new XAttribute( XName.Get( "salt", w.NamespaceName ), Convert.ToBase64String( encryptionObj.SaltArray ) ) );

      _settings.Root.AddFirst( documentProtection );
    }

    public void RemovePasswordProtection( string password )
    {
      if( !IsPasswordProtected || string.IsNullOrEmpty( password ) )
        return;

      var documentProtection = _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).FirstOrDefault();

      if( documentProtection != null )
      {
        var hash = documentProtection.Attribute( XName.Get( "hash", w.NamespaceName ) );

        if( hash != null )
        {
          var keyValues = ComputeHashValue( password, documentProtection );

          if( !string.IsNullOrEmpty( keyValues ) && hash.Value != keyValues )
          {
            throw new UnauthorizedAccessException( "Invalid password." );
          }
        }
      }

      _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).Remove();
    }

    public void SetDefaultFont( Font fontFamily, double fontSize = 11d, Xceed.Drawing.Color? fontColor = null )
    {
      var docDefault = this.GetDocDefaults();
      if( docDefault != null )
      {
        var rPrDefault = docDefault.Element( XName.Get( "rPrDefault", w.NamespaceName ) );
        if( rPrDefault != null )
        {
          var rPr = rPrDefault.Element( XName.Get( "rPr", w.NamespaceName ) );
          if( rPr != null )
          {
            //rFonts
            if( fontFamily != null )
            {
              var rFonts = rPr.Element( XName.Get( "rFonts", w.NamespaceName ) );
              if( rFonts == null )
              {
                rPr.AddFirst( new XElement( XName.Get( "rFonts", w.NamespaceName ) ) );
                rFonts = rPr.Element( XName.Get( "rFonts", w.NamespaceName ) );
              }

              rFonts.Attributes().Remove();
              rFonts.SetAttributeValue( XName.Get( "ascii", w.NamespaceName ), fontFamily.Name );
              rFonts.SetAttributeValue( XName.Get( "hAnsi", w.NamespaceName ), fontFamily.Name );
              rFonts.SetAttributeValue( XName.Get( "cs", w.NamespaceName ), fontFamily.Name );
              rFonts.SetAttributeValue( XName.Get( "eastAsia", w.NamespaceName ), fontFamily.Name );
            }

            //sz
            var sz = rPr.Element( XName.Get( "sz", w.NamespaceName ) );
            if( sz == null )
            {
              rPr.Add( new XElement( XName.Get( "sz", w.NamespaceName ) ) );
              sz = rPr.Element( XName.Get( "sz", w.NamespaceName ) );
            }
            sz.SetAttributeValue( XName.Get( "val", w.NamespaceName ), fontSize * 2 );

            //szCs
            var szCs = rPr.Element( XName.Get( "szCs", w.NamespaceName ) );
            if( szCs == null )
            {
              rPr.Add( new XElement( XName.Get( "szCs", w.NamespaceName ) ) );
              szCs = rPr.Element( XName.Get( "szCs", w.NamespaceName ) );
            }
            szCs.SetAttributeValue( XName.Get( "val", w.NamespaceName ), fontSize * 2 );

            //color
            if( ( fontColor != null ) && fontColor.HasValue )
            {
              var color = rPr.Element( XName.Get( "color", w.NamespaceName ) );
              if( color == null )
              {
                rPr.Add( new XElement( XName.Get( "color", w.NamespaceName ) ) );
                color = rPr.Element( XName.Get( "color", w.NamespaceName ) );
              }
              color.SetAttributeValue( XName.Get( "val", w.NamespaceName ), fontColor.Value.ToHex() );
            }
          }
        }
      }
    }



















    public T AddChart<T>() where T : Chart, new()
    {
      var chart = new T();
      chart.PackagePart = this.CreateChartPackagePart();
      chart.RelationPackage = this.CreateChartRelationShip( chart.PackagePart );
      return chart;
    }
    #endregion

    #region Internal Methods

    protected internal virtual Xceed.Document.NET.Document InternalCopy( bool closePackage = false )
    {
      return null;
    }

    protected internal virtual void SaveHeadersFooters()
    {
    }






    internal XElement CreateEdit( EditType t, DateTime edit_time, object content )
    {
      if( t == EditType.del )
      {
        foreach( object o in (IEnumerable<XElement>)content )
        {
          if( o is XElement )
          {
            XElement e = ( o as XElement );
            IEnumerable<XElement> ts = e.DescendantsAndSelf( XName.Get( "t", Document.w.NamespaceName ) );

            for( int i = 0; i < ts.Count(); i++ )
            {
              XElement text = ts.ElementAt( i );
              text.ReplaceWith( new XElement( Document.w + "delText", text.Attributes(), text.Value ) );
            }
          }
        }
      }

      // Check the author in a Try/Catch 
      // (for the cases where we do not have the rights to access that information)
      string author = "";
      try
      {
        author = Environment.UserDomainName + "\\" + Environment.UserName;
      }
      catch( Exception )
      {
        // do nothing
      }

      if( _lastEditId < 0 )
      {
        var trackerIDs = ( from d in _mainDoc.Descendants()
                        where d.Name.LocalName == "ins" || d.Name.LocalName == "del"
                        select d.Attribute( XName.Get( "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main" ) ) );

        if( trackerIDs.Count() > 0 )
        {
          _lastEditId = trackerIDs.Select( id => int.Parse( id.Value ) ).Max();
        }
      }

      if( author.Trim() == "" )
      {
        return
        (
            new XElement( Document.w + t.ToString(),
                new XAttribute( Document.w + "id", ++_lastEditId ),
                new XAttribute( Document.w + "date", edit_time ),
            content )
        );
      }

      return
      (
          new XElement( Document.w + t.ToString(),
              new XAttribute( Document.w + "id", ++_lastEditId ),
              new XAttribute( Document.w + "author", author ),
              new XAttribute( Document.w + "date", edit_time ),
          content )
      );
    }

    internal XDocument GetSettings()
    {
      if( _settings != null )
        return _settings;

      XDocument settings;
      using( TextReader tr = new StreamReader( _settingsPart.GetStream() ) )
      {
        settings = XDocument.Load( tr );
      }

      return settings ?? null;
    }

    internal XDocument GetFootnotes()
    {
      return _footnotes;
    }

    internal XDocument GetEndnotes()
    {
      return _endnotes;
    }

    protected internal override void AddElementInXml( object element )
    {
      // Add element just before the last sectPr.
      var body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
      var sectPr = body.Elements( XName.Get( "sectPr", w.NamespaceName ) ).Last();

      sectPr.AddBeforeSelf( element );
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
      var mainDocPart = ( ( documentType == DocumentTypes.Document ) || ( documentType == DocumentTypes.Pdf ) )
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

    internal static Document PostLoad( ref Package package, Document document, DocumentTypes documentType )
    {
      //var document = (documentType == DocumentTypes.Pdf) ? new DocPDF( null, null ) as Document : new DocX( null, null );
      document._package = package;
      document.Document = document;

      #region MainDocumentPart
      document.PackagePart = HelperFunctions.GetMainDocumentPart( package );

      using( TextReader tr = new StreamReader( document.PackagePart.GetStream( FileMode.Open, FileAccess.Read ) ) )
      {
        document._mainDoc = XDocument.Load( tr, LoadOptions.PreserveWhitespace );
      }
      #endregion

      Document.PopulateDocument( document, package );

      using( TextReader tr = new StreamReader( document._settingsPart.GetStream() ) )
      {
        document._settings = XDocument.Load( tr );
      }

      document._paragraphLookup.Clear();
      var paragraphs = document.Paragraphs;
      foreach( var p in paragraphs )
      {
        if( !document._paragraphLookup.ContainsKey( p.EndIndex ) )
        {
          document._paragraphLookup.Add( p.EndIndex, p );
        }
      }

      var docType = document._mainDoc.DocumentType;
      if( ( docType != null ) && !string.IsNullOrEmpty( docType.InternalSubset ) )
      {
        if( docType.InternalSubset.Contains( "ENTITY" ) )
          throw new InvalidDataException( "Document to load contains external entities which are not supported" );
      }

      return document;
    }

    internal static Document Load( Stream stream, Document document, DocumentTypes documentType )
    {
      var ms = new MemoryStream();

      try
      {
        stream.Position = 0;
      }
      catch( Exception )
      {
        // no stream.Position for Web streams.
      }

      HelperFunctions.CopyStream( stream, ms );

      // Open the docx package
      var package = Package.Open( ms, FileMode.Open, FileAccess.ReadWrite );

      document = Document.PostLoad( ref package, document, documentType );
      document._package = package;
      document._memoryStream = ms;
      document._stream = stream;

      return document;
    }

    internal static Document Load( string filename, Document document, DocumentTypes documentType )
    {
      var ms = new MemoryStream();

      if( File.Exists( filename ) )
      {
        using( FileStream fs = new FileStream( filename, FileMode.Open, FileAccess.Read, FileShare.Read ) )
        {
          HelperFunctions.CopyStream( fs, ms );
        }
      }
      else
      {
        WebRequest request = null;
        HttpWebResponse response = null;
        Stream receiveStream = null;
        try
        {
          ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
          request = (HttpWebRequest)WebRequest.Create( filename );
          response = (HttpWebResponse)request.GetResponse();
          receiveStream = response.GetResponseStream();
          HelperFunctions.CopyStream( receiveStream, ms );
        }
        catch( Exception )
        {
          throw new FileNotFoundException( string.Format( "File could not be found {0}", filename ) );
        }
        finally
        {
          if( response != null )
          {
            response.Close();
          }
          if( receiveStream != null )
          {
            receiveStream.Close();
          }
        }
      }

      // Open the docx package
      var package = Package.Open( ms, FileMode.Open, FileAccess.ReadWrite );

      document = PostLoad( ref package, document, documentType );
      document._package = package;
      document.SetFileName( filename );
      document._memoryStream = ms;

      return document;
    }

    internal void SetFileName( string filename )
    {
      _filename = ( filename != null )
                  ? filename.EndsWith( ".docx" ) || filename.EndsWith( ".doc" ) || filename.EndsWith( ".docm" ) ? filename : filename + ".docx"
                  : null;
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

    internal Image AddImage( object o, string contentType = "image/jpeg" )
    {
      // Open a Stream to the new image being added.
      var newImageStream = ( o is string ) ? new FileStream( o as string, FileMode.Open, FileAccess.Read ) : o as Stream;
      using( newImageStream )
      {

        //all document's parts, including image parts.
        //var partLookup = _package.GetParts().ToDictionary( x => x.Uri.ToString(), x => x, StringComparer.Ordinal );

        //var imageParts = new List<PackagePart>();
        //var relationshipImages = this.PackagePart.GetRelationshipsByType( RelationshipImage );
        //foreach( var item in relationshipImages )
        //{
        //  var targetUri = item.TargetUri.ToString();
        //  PackagePart part;
        //  if( partLookup.TryGetValue( targetUri, out part ) )
        //  {
        //    // all document's used image parts.
        //    imageParts.Add( part );
        //  }
        //}

        //var relsParts = partLookup
        // .Where(
        //   item =>
        //   item.Value.ContentType.Equals( ContentTypeApplicationRelationShipXml, StringComparison.Ordinal ) &&
        //   item.Key.IndexOf( "/word/", StringComparison.Ordinal ) > -1 )
        // .Select( item => item.Value );

        //var xNameTarget = XName.Get( "Target" );
        //var xNameTargetMode = XName.Get( "TargetMode" );

        //foreach( var relsPart in relsParts )
        //{
        //  XDocument relsPartContent;
        //  using( var tr = new StreamReader( relsPart.GetStream( FileMode.Open, FileAccess.Read ) ) )
        //  {
        //    relsPartContent = XDocument.Load( tr );
        //  }

        //  // relationship parts of images.
        //  var imageRelationships =
        //  relsPartContent.Root.Elements().Where
        //  (
        //      imageRel =>
        //      imageRel.Attribute( XName.Get( "Type" ) ).Value.Equals( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" )
        //  );

        //  foreach( var imageRelationship in imageRelationships )
        //  {
        //    var attribute = imageRelationship.Attribute( xNameTarget );
        //    if( attribute != null )
        //    {
        //      var targetModeAttr = imageRelationship.Attribute( xNameTargetMode );
        //      var targetMode = ( targetModeAttr != null ) ? targetModeAttr.Value : string.Empty;

        //      if( !targetMode.Equals( "External" ) )
        //      {
        //        var imagePartUri = Path.Combine( Path.GetDirectoryName( relsPart.Uri.ToString() ), attribute.Value );
        //        imagePartUri = Path.GetFullPath( imagePartUri.Replace( "\\_rels", string.Empty ) );
        //        imagePartUri = imagePartUri.Replace( Path.GetFullPath( "\\" ), string.Empty ).Replace( "\\", "/" );

        //        if( !imagePartUri.StartsWith( "/" ) )
        //        {
        //          imagePartUri = "/" + imagePartUri;
        //        }

        //        var imagePart = _package.GetPart( new Uri( imagePartUri, UriKind.Relative ) );
        //        imageParts.Add( imagePart );
        //      }
        //    }
        //  }
        //}

        //foreach( var pp in imageParts )
        //{
        //  // Get the image object for this image part.
        //  using( var tempStream = pp.GetStream( FileMode.Open, FileAccess.Read ) )
        //  {
        //    // Compare this image to the new image being added.
        //    if( HelperFunctions.IsSameFile( tempStream, newImageStream ) )
        //    {
        //      // Return the Image object
        //      var relationship = relationshipImages.First( x => x.TargetUri == pp.Uri );
        //      return new Image( this, relationship );
        //    }
        //  }
        //}

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
          HelperFunctions.CopyStream( newImageStream, stream, bufferSize: 4096 );
        }// Close the Stream to the new image part.

        return new Image( this, rel );
      }
    }

    internal Image GetImage( Stream stream )
    {
      if( stream == null )
        return null;

      stream.Position = 0;

      var imageParts = new Dictionary<PackageRelationship, PackagePart>();
      var partLookup = _package.GetParts().ToDictionary( x => x.Uri.ToString(), x => x, StringComparer.Ordinal );

      // all image relationships.
      var relationshipImages = this.PackagePart.GetRelationshipsByType( RelationshipImage );
      // take all used images (from relationships)
      foreach( var item in relationshipImages )
      {
        var targetUri = item.TargetUri.ToString();
        PackagePart part;
        if( partLookup.TryGetValue( targetUri, out part ) )
        {
          // all document's used image parts.
          imageParts.Add( item, part );
        }
      }

      foreach( var entry in imageParts )
      {
        // Get the image object for this image part.
        using( var tempStream = entry.Value.GetStream( FileMode.Open, FileAccess.Read ) )
        {
          // Compare this image to the new image being added.
          if( HelperFunctions.IsSameFile( tempStream, stream ) )
            return new Image( this, entry.Key );
        }
      }

      return null;
    }

    internal static void UpdateCorePropertyValue( Document document, string corePropertyName, string corePropertyValue )
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
          Xceed.Document.NET.Text.PreserveSpace( t );
          e.Add( new XElement( firstRun.Name, firstRun.Attributes(), firstRun.Element( XName.Get( "rPr", w.NamespaceName ) ), t ) );
        }
      }

      // A list of documents, which will contain, if they exist: header1, header2, header3, footer1, footer2, footer3.
      var documents = new List<XElement> { };

      foreach( var section in document.Sections )
      {
        // Headers
        var headers = section.Headers;
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

        // Footers
        var footers = section.Footers;
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
      }

      foreach( var doc in documents )
      {
        foreach( XElement e in doc.Descendants( XName.Get( "fldSimple", w.NamespaceName ) ) )
        {
          string attr_value = e.Attribute( XName.Get( "instr", w.NamespaceName ) ).Value.Replace( " ", string.Empty ).Trim().ToLower();
          if( Regex.IsMatch( attr_value, matchPattern ) )
          {
            var firstRun = e.Element( w + "r" );

            // Delete everything and insert updated text value
            e.RemoveNodes();

            var t = new XElement( w + "t", corePropertyValue );
            Xceed.Document.NET.Text.PreserveSpace( t );
            e.Add( new XElement( firstRun.Name, firstRun.Attributes(), firstRun.Element( XName.Get( "rPr", w.NamespaceName ) ), t ) );
          }
        }
      }

      document.SaveHeadersFooters();

      // The following line was commnented in the regard to case 2744 
      // Document.PopulateDocument( document, document._package );
    }

    internal static void UpdateCustomPropertyValue( Document document, CustomProperty cp )
    {
      var customPropertyName = cp.Name;
      var customPropertyValue = ( cp.Value ?? "" ).ToString();

      // A list of documents, which will contain, The Main Document and if they exist: header1, header2, header3, footer1, footer2, footer3.
      var documents = new List<XElement> { document._mainDoc.Root };

      // Check if each header exists and add it if so.
      #region Headers
      foreach( var section in document.Sections )
      {
        var headers = section.Headers;
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
      }
      #endregion

      // Check if each footer exists and add if if so.
      #region Footers
      foreach( var section in document.Sections )
      {
        var footers = section.Footers;
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
      }
      #endregion

      var match_value = string.Format( @"DOCPROPERTY  {0}  \* MERGEFORMAT", customPropertyName.Contains( " " ) ? "\"" + customPropertyName + "\"" : customPropertyName )
                              .Replace( " ", string.Empty );

      // Process each document in the list.
      foreach( XElement doc in documents )
      {
        #region Word 2010+
        var instrTexts = doc.Descendants( XName.Get( "instrText", w.NamespaceName ) );
        for( int i = 0; i < instrTexts.Count(); ++i )
        {
          var e = instrTexts.ElementAt( i );
          var attr_value = e.Value.Replace( " ", string.Empty ).Trim();

          if( attr_value.StartsWith( "DOCPROPERTY" ) && !attr_value.EndsWith( "MERGEFORMAT" ) && ( i < instrTexts.Count() - 1 ) )
          {
            var nextInstrText = instrTexts.ElementAt( i + 1 );
            var nextAttr_value = nextInstrText.Value.Replace( " ", string.Empty ).Trim();
            if( !nextAttr_value.StartsWith( "DOCPROP" ) && nextAttr_value.EndsWith( "FORMAT" ) )
            {
              attr_value += nextAttr_value;
              i++;
            }
          }

          if( attr_value.Equals( match_value, StringComparison.CurrentCultureIgnoreCase ) )
          {
            var node = e.Parent.NextNode;
            bool found = false;

            while( node != null )
            {
              if( node.NodeType == XmlNodeType.Element )
              {
                var ele = node as XElement;

                var match = ele.Descendants( XName.Get( "t", w.NamespaceName ) );
                if( match.Any() )
                {
                  if( !found )
                  {
                    var newValue = HelperFunctions.FormatInput( customPropertyValue, cp.Formatting?.Xml );

                    if( ele.Name.LocalName == "r" ) // Main run
                    {
                      var runElements = ele.Elements().Where( el => el.Name.LocalName != "t" ).ToList();
                      var nextElement = ele.ElementsAfterSelf().FirstOrDefault();

                      var lastRunTag = newValue.LastOrDefault();

                      if( lastRunTag != null )
                      {
                        var targetElement = lastRunTag.Descendants( XName.Get( "t", w.NamespaceName ) ).FirstOrDefault();

                        if( targetElement != null )
                        {
                          targetElement.AddBeforeSelf( runElements );
                        }
                      }

                      ele.AddBeforeSelf( newValue );
                      ele.Remove();
                      node = nextElement; // Move to the next element
                      continue;
                    }

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

              // Move to the next node
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
            Xceed.Document.NET.Text.PreserveSpace( t );
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

    internal XElement GetDocDefaults()
    {
      if( _styles == null )
        return null;

      return _styles.Element( XName.Get( "styles", Document.w.NamespaceName ) ).Element( XName.Get( "docDefaults", Document.w.NamespaceName ) );
    }

    internal long GetNextFreeDocPrId()
    {
      lock( nextFreeDocPrIdLock )
      {
        if( nextFreeDocPrId != null )
        {
          nextFreeDocPrId++;
          return nextFreeDocPrId.Value;
        }

        long newDocPrId = 1;
        var existingIds = new HashSet<string>();
        //In Body
        existingIds.UnionWith( this.GetExistingDocPrIds( this.Xml ) );

        foreach( var section in this.Sections )
        {
          // In Headers.
          var headers = section.Headers;
          if( headers != null )
          {
            if( headers.Odd != null )
            {
              existingIds.UnionWith( this.GetExistingDocPrIds( headers.Odd.Xml ) );
            }
            if( headers.Even != null )
            {
              existingIds.UnionWith( this.GetExistingDocPrIds( headers.Even.Xml ) );
            }
            if( headers.First != null )
            {
              existingIds.UnionWith( this.GetExistingDocPrIds( headers.First.Xml ) );
            }
          }

          // In Footers.
          var footers = section.Footers;
          if( footers != null )
          {
            if( footers.Odd != null )
            {
              existingIds.UnionWith( this.GetExistingDocPrIds( footers.Odd.Xml ) );
            }

            if( footers.Even != null )
            {
              existingIds.UnionWith( this.GetExistingDocPrIds( footers.Even.Xml ) );
            }

            if( footers.First != null )
            {
              existingIds.UnionWith( this.GetExistingDocPrIds( footers.First.Xml ) );
            }
          }
        }

        if( existingIds.Count > 0 )
        {
          newDocPrId = existingIds.Max( id => long.Parse( id ) ) + 1;
        }

        nextFreeDocPrId = newDocPrId;
        return nextFreeDocPrId.Value;
      }
    }

    internal string GetNormalStyleId()
    {
      if( _defaultParagraphStyleId != null )
        return _defaultParagraphStyleId;

      var normalStyle =
      (
          from s in _styles.Element( XName.Get( "styles", Document.w.NamespaceName ) ).Elements()
          let type = s.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
          let def = s.Attribute( XName.Get( "default", Document.w.NamespaceName ) )
          where ( ( type != null ) && ( type.Value == "paragraph" ) ) && ( ( def != null ) && ( def.Value == "1" ) )
          select s
      ).FirstOrDefault();

      if( normalStyle != null )
      {
        var styleId = normalStyle.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) );
        if( styleId != null )
        {
          _defaultParagraphStyleId = styleId.Value;
        }
      }

      if( _defaultParagraphStyleId == null )
      {
        _defaultParagraphStyleId = "Normal";
      }

      return _defaultParagraphStyleId;
    }

    internal static void PrepareDocument( ref Document document, DocumentTypes documentType )
    {
      // Store this document in memory
      var ms = new MemoryStream();

      // Create the docx package
      var package = Package.Open( ms, FileMode.Create, FileAccess.ReadWrite );

      Document.PostCreation( package, documentType );
      document = Document.Load( ms, document, documentType );
    }

    internal void UpdateCacheSections()
    {
      // Save Headers/Footers in case they were modified. GetSectionCaches() will create new Sections based on them.
      this.SaveHeadersFooters();

      this.ClearParagraphsCache();
      _cachedSections = this.GetSections();
    }

    internal Paragraph CreateChartElement( Chart chart, Paragraph paragraph, float width, float height )
    {

      if( chart.PackagePart == null )
      {
        chart.PackagePart = this.CreateChartPackagePart();
        chart.RelationPackage = this.CreateChartRelationShip( chart.PackagePart );
      }
      var relID = chart.RelationPackage.Id.ToString();

      // Save a chart info the chartPackagePart
      var p = ( paragraph == null ) ? InsertParagraph() : paragraph;

      using( TextWriter tw = new StreamWriter( chart.PackagePart.GetStream( FileMode.Create, FileAccess.Write ) ) )
      {
        chart.ExternalXml.Save( tw );
      }

      // Insert a new chart into a paragraph.
      chart.SetInternalChartSettings( relID, width, height );

      return p;
    }

    internal PackagePart CreateChartPackagePart()
    {
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


      return chartPackagePart;
    }

    internal PackageRelationship CreateChartRelationShip( PackagePart chartPackagePart )
    {
      // Create a new chart relationship
      var relID = this.GetNextFreeRelationshipID();
      return this.PackagePart.CreateRelationship( chartPackagePart.Uri, TargetMode.Internal, Document.RelationshipChart, relID );
    }

    internal double GetDocDefaultFontSize()
    {
      var docDefault = this.Document.GetDocDefaults();
      if( docDefault != null )
      {
        var rPrDefault = docDefault.Element( XName.Get( "rPrDefault", Document.w.NamespaceName ) );
        if( rPrDefault != null )
        {
          var rPr = rPrDefault.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
          if( rPr != null )
          {
            var size = rPr.Element( XName.Get( "sz", Document.w.NamespaceName ) );
            if( size != null )
              return Convert.ToDouble( size.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) ) / 2;
          }
        }
      }

      return Document.DefaultFontSize;
    }

    internal string GetDocDefaultFontFamily()
    {
      var docDefault = this.GetDocDefaults();
      if( docDefault != null )
      {
        var rPrDefault = docDefault.Element( XName.Get( "rPrDefault", Document.w.NamespaceName ) );
        if( rPrDefault != null )
        {
          var rPr = rPrDefault.Element( XName.Get( "rPr", Document.w.NamespaceName ) );
          if( rPr != null )
          {
            var rFonts = rPr.Element( XName.Get( "rFonts", Document.w.NamespaceName ) );
            if( rFonts != null )
            {
              var fontName = rFonts.GetAttribute( XName.Get( "ascii", Document.w.NamespaceName ) )
                            ?? rFonts.GetAttribute( XName.Get( "hAnsi", Document.w.NamespaceName ) )
                            ?? rFonts.GetAttribute( XName.Get( "cs", Document.w.NamespaceName ) )
                            ?? rFonts.GetAttribute( XName.Get( "hint", Document.w.NamespaceName ) )
                            ?? rFonts.GetAttribute( XName.Get( "eastAsia", Document.w.NamespaceName ) );
              if( !string.IsNullOrEmpty( fontName ) )
                return fontName;
            }
          }
        }
      }

      return  "Calibri";
    }

    internal static string ComputeHashValue( string password, XElement documentProtection )
    {
      var salt = documentProtection.Attribute( XName.Get( "salt", w.NamespaceName ) );
      var cryptProviderType = documentProtection.Attribute( XName.Get( "cryptProviderType", w.NamespaceName ) );
      var cryptAlgorithmSid = documentProtection.Attribute( XName.Get( "cryptAlgorithmSid", w.NamespaceName ) );

      var keyValues = "";

      if( salt != null && cryptProviderType != null && cryptAlgorithmSid != null )
      {
        var saltValue = salt.Value;
        var cryptProviderTypeValue = cryptProviderType.Value;
        var cryptAlgorithmSidValue = cryptAlgorithmSid.Value;

        if( !string.IsNullOrEmpty( saltValue ) && !string.IsNullOrEmpty( cryptProviderTypeValue ) )
        {
          HashAlgorithmName hashAlgorithm = HashAlgorithmName.NotSupported;

          switch( cryptProviderTypeValue )
          {
            case "rsaFull":
              switch( cryptAlgorithmSidValue )
              {
                case "1":
                case "3":
                  hashAlgorithm = HashAlgorithmName.SHA1;
                  break;
                case "4":
                  hashAlgorithm = HashAlgorithmName.SHA1Managed;
                  break;
              }
              break;
            case "rsaAES":
              switch( cryptAlgorithmSidValue )
              {
                case "12":
                  hashAlgorithm = HashAlgorithmName.SHA256Managed;
                  break;
                case "13":
                  hashAlgorithm = HashAlgorithmName.SHA384Managed;
                  break;
                case "14":
                  hashAlgorithm = HashAlgorithmName.SHA512Managed;
                  break;
              }
              break;
          }

          _hashAlgorithmName = hashAlgorithm;
          var encryption = new Encryption();
          keyValues = encryption.Decrypt( password, saltValue, hashAlgorithm );
        }
      }

      return keyValues;
    }

    #endregion

    #region Private Methods

















    private void MergeDocumentsPackages( Document remote_document, MergingMode mergingMode, out XDocument remote_mainDoc, out XElement remote_body, out XElement local_body, out PackagePartCollection ppc )
    {
      // We don't want to affect the original XDocument, so create a new one from the old one.
      remote_mainDoc = new XDocument( remote_document._mainDoc );
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

      // Get the body of the remote document.
      remote_body = remote_mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );
      // Get the body of the local document.
      local_body = _mainDoc.Root.Element( XName.Get( "body", w.NamespaceName ) );

     // Remove all header and footer references.
      remote_mainDoc.Descendants( XName.Get( "headerReference", w.NamespaceName ) ).Remove();
      remote_mainDoc.Descendants( XName.Get( "footerReference", w.NamespaceName ) ).Remove();

      // Every file that is missing from the local document will have to be copied, every file that already exists will have to be merged.
      ppc = remote_document._package.GetParts();
      var ignoreContentTypes = new List<string>
            {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",                
                "application/vnd.openxmlformats-package.core-properties+xml",
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                ContentTypeApplicationRelationShipXml
            };

      // Check if each PackagePart pp exists in this document.
      foreach( PackagePart remote_pp in ppc )
      {
        if( ignoreContentTypes.Contains( remote_pp.ContentType ) || _imageContentTypes.Contains( remote_pp.ContentType ) )
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
              merge_footnotes( remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes );
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
              merge_endnotes( remote_pp, local_pp, remote_mainDoc, remote_document, remote_endnotes );
              break;

            case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
              merge_styles( remote_pp, local_pp, remote_mainDoc, remote_document, _footnotes, _endnotes, mergingMode );
              break;

            // Merges Styles after merging the footnotes, so the changes will be applied to the correct document/footnotes.
            case "application/vnd.ms-word.stylesWithEffects+xml":
              merge_styles( remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes, remote_endnotes, mergingMode );
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

    }

    private void MergeDocumentsRelations( Document remote_document, XDocument remote_mainDoc, XElement remote_body, PackagePartCollection ppc )
    {
      foreach( var hyperlink_rel in remote_document.PackagePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ) )
      {
        var old_rel_Id = hyperlink_rel.Id;
        var new_rel_Id = this.PackagePart.CreateRelationship( hyperlink_rel.TargetUri, hyperlink_rel.TargetMode, hyperlink_rel.RelationshipType ).Id;
        var hyperlink_refs = remote_mainDoc.Descendants( XName.Get( "hyperlink", w.NamespaceName ) );
        var hyperlink_refs2 = remote_mainDoc.Descendants( XName.Get( "hlinkClick", a.NamespaceName ) );
        var links = hyperlink_refs.Concat( hyperlink_refs2 );
        foreach( var hyperlink_ref in links )
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

      var remoteNumberingRelationship = remote_document.PackagePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" );
      var localNumberingRelationship = this.PackagePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" );
      if( ( remoteNumberingRelationship.Count() > 0 ) && ( localNumberingRelationship.Count() == 0 ) )
      {
        foreach( var rel in remoteNumberingRelationship )
        {
          this.PackagePart.CreateRelationship( rel.TargetUri, rel.TargetMode, rel.RelationshipType );
        }
      }

      //var remoteCustomXmlRelationship = remote_document.PackagePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" );
      //var localCustomXmlRelationship = this.PackagePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" );
      //if( ( remoteCustomXmlRelationship.Count() > 0 ) && ( localCustomXmlRelationship.Count() == 0 ) )
      //{
      //  foreach( var rel in remoteCustomXmlRelationship )
      //  {
      //    var uri = new Uri( "../" + rel.TargetUri.OriginalString, UriKind.Relative);
      //    this.PackagePart.CreateRelationship( uri, rel.TargetMode, rel.RelationshipType );
      //  }
      //}

      if( ( remote_document._fontTablePart != null ) && ( this._fontTablePart != null ) )
      {
        var remoteFontRelationship = remote_document._fontTablePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" );
        var localFontRelationship = this._fontTablePart.GetRelationshipsByType( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" );
        if( ( remoteFontRelationship.Count() > 0 ) && ( localFontRelationship.Count() == 0 ) )
        {
          foreach( var rel in remoteFontRelationship )
          {
            this._fontTablePart.CreateRelationship( rel.TargetUri, rel.TargetMode, rel.RelationshipType );
          }
        }
      }

      foreach( PackagePart remote_pp in ppc )
      {
        if( _imageContentTypes.Contains( remote_pp.ContentType ) )
        {
          this.MergeImages( remote_pp, remote_document, remote_mainDoc, remote_pp.ContentType );
        }

        if( remote_pp.Uri.OriginalString.StartsWith( "/word/charts/" ) )
        {
          this.MergeCharts( remote_pp, remote_document, remote_mainDoc );
        }
      }

      int id = 0;
      var local_docPrs = _mainDoc.Root.Descendants( XName.Get( "docPr", wp.NamespaceName ) );
      foreach( var local_docPr in local_docPrs )
      {
        var a_id = local_docPr.Attribute( XName.Get( "id" ) );
        int a_id_value;
        if( a_id != null && HelperFunctions.TryParseInt( a_id.Value, out a_id_value ) )
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
    }

    private void AddRootAttributes( XDocument remote_mainDoc )
    {
      // Copy any missing root attributes to the local document.
      foreach( XAttribute a in remote_mainDoc.Root.Attributes() )
      {
        if( _mainDoc.Root.Attribute( a.Name ) == null )
        {
          _mainDoc.Root.SetAttributeValue( a.Name, a.Value );
        }
      }
    }

    private void DeleteHeadersOrFooters( bool isHeader, bool deleteReferences = true )
    {
      var reference = isHeader ? "header" : "footer";

      // Get all header Relationships in this document.
      var header_relationships = this.PackagePart.GetRelationshipsByType( string.Format( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference ) );

      foreach( var header_relationship in header_relationships )
      {
        var body = _mainDoc.Descendants( XName.Get( "body", w.NamespaceName ) ).FirstOrDefault();
        if( body != null )
        {
          this.DeleteHeaderOrFooter( header_relationship, reference, _package, body, deleteReferences );
        }
      }
    }

    private void DeleteHeaderOrFooter( PackageRelationship header_relationship, string reference, Package package, XElement documentBody, bool deleteReferences = true )
    {
      // Get the TargetUri for this Part.
      Uri header_uri = header_relationship.TargetUri;

      // Check to see if the document actually contains the Part.
      if( !header_uri.OriginalString.StartsWith( "/word/" ) )
        header_uri = new Uri( "/word/" + header_uri.OriginalString, UriKind.Relative );

      if( package.PartExists( header_uri ) )
      {
        // Delete the Part
        package.DeletePart( header_uri );

        if( deleteReferences )
        {
          // Remove all references to this Relationship in the document.
          this.RemoveReferences( header_relationship, reference, documentBody );
        }

        // Delete the Relationship.
        package.DeleteRelationship( header_relationship.Id );
      }
    }

    private void RemoveReferences( PackageRelationship header_relationship, string reference, XElement documentBody )
    {
      // Get all references to this Relationship in the document.
      var query =
      (
          from e in documentBody.Descendants()
          where ( e.Name.LocalName == string.Format( "{0}Reference", reference ) ) && ( e.Attribute( r + "id" ).Value == header_relationship.Id )
          select e
      );

      // Remove all references to this Relationship in the document.
      for( int i = 0; i < query.Count(); i++ )
      {
        query.ElementAt( i ).Remove();
      }
    }

    internal Paragraph InsertParagraphPictures( Paragraph p )
    {
      if( ( p != null ) && ( p.Pictures != null ) && ( p.Pictures.Count > 0 ) )
      {
        var clonedParagraph = new Paragraph( p.Xml );
        clonedParagraph.Document = p.Document.Copy();
        clonedParagraph.PackagePart = p.PackagePart;

        var ppc = clonedParagraph.Document._package.GetParts();

        foreach( var remote_pp in ppc )
        {
          if( _imageContentTypes.Contains( remote_pp.ContentType ) )
          {
            // This will modify the clonedParagraph's Picture Ids before adding it in this Document with the following base.InsertParagraph( updatedParagraph );
            this.MergeImages( remote_pp, clonedParagraph.Document, clonedParagraph.Document._mainDoc, remote_pp.ContentType );
          }
        }

        // return an updated paragraph(new image Ids) of the p paragraph.
        return clonedParagraph.Document.Paragraphs.First( paragraph => paragraph.StartIndex == p.StartIndex );
      }

      return p;
    }

    internal void InsertParagraphStyles( Paragraph p )
    {
      // Check if style is already present in document.
      if( ( p != null ) && ( p.StyleId != this.GetNormalStyleId() ) )
      {
        var styleToFind = HelperFunctions.GetStyle( this, p.StyleId );

        // style is not there, we need to add it.
        if( styleToFind == null )
        {
          var styleIdToFind = p.StyleId;
          while( ( styleIdToFind != null ) && ( styleIdToFind != this.GetNormalStyleId() ) )
          {
            var styleToAdd = HelperFunctions.GetStyle( p.Document, styleIdToFind );
            if( styleToAdd != null )
            {
              _styles.Root.Add( styleToAdd );

              var basedOn = styleToAdd.Element( XName.Get( "basedOn", Document.w.NamespaceName ) );
              styleIdToFind = ( basedOn != null ) ? basedOn.Attribute( XName.Get( "val", Document.w.NamespaceName ) ).Value : null;
            }
          }
        }
      }
    }

    private void MergeImages( PackagePart remote_pp, Document remote_document, XDocument remote_mainDoc, String contentType )
    {
      // Before doing any other work, check to see if this image is actually referenced in the document.
      // In my testing I have found cases of Images inside documents that are not referenced
      var remote_rel = remote_document.PackagePart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( remote_pp.Uri.OriginalString.Replace( "/word/", "" ) ) ).FirstOrDefault();
      if( remote_rel == null )
      {
        remote_rel = remote_document.PackagePart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( remote_pp.Uri.OriginalString ) ).FirstOrDefault();
        if( remote_rel == null )
        {
          if( remote_document._numberingPart == null )
            return;

          // Look for images in _numbering.
          remote_rel = remote_document._numberingPart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( remote_pp.Uri.OriginalString.Replace( "/word/", "" ) ) ).FirstOrDefault();
          if( remote_rel == null )
          {
            remote_rel = remote_document._numberingPart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( remote_pp.Uri.OriginalString ) ).FirstOrDefault();
            if( remote_rel == null )
              return;
          }
        }
      }

      this.MergeImagesCore( remote_rel, remote_pp, remote_mainDoc, contentType, this.PackagePart );
    }

    private void MergeImagesCore( PackageRelationship relationship, PackagePart remote_pp, XDocument remote_mainDoc, String contentType, PackagePart packagePartToSaveTo )
    {
      var remote_Id = relationship.Id;
      bool found = false;

      using( Stream s_read = remote_pp.GetStream() )
      {
        var remote_hash = this.ComputeMD5HashString( s_read );
        var image_parts = _package.GetParts().Where( pp => pp.ContentType.Equals( contentType ) );

        foreach( var part in image_parts )
        {
          using( Stream partStream = part.GetStream() )
          {
            var local_hash = ComputeMD5HashString( partStream );
            if( local_hash.Equals( remote_hash ) )
            {
              // This image already exists in this document.
              found = true;

              var local_rel = packagePartToSaveTo.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( part.Uri.OriginalString.Replace( "/word/", "" ) ) ).FirstOrDefault();
              if( local_rel == null )
              {
                local_rel = packagePartToSaveTo.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( part.Uri.OriginalString ) ).FirstOrDefault();
                if( local_rel == null )
                {
                  // Look in _numbering.
                  if( _numberingPart != null )
                  {
                    local_rel = _numberingPart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( part.Uri.OriginalString.Replace( "/word/", "" ) ) ).FirstOrDefault();
                    if( local_rel == null )
                    {
                      local_rel = _numberingPart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( part.Uri.OriginalString ) ).FirstOrDefault();
                    }
                  }
                }
              }

              if( local_rel != null )
              {
                var new_Id = local_rel.Id;

                // Replace all instances of remote_Id in the local document with local_Id
                this.ReplaceAllRemoteID( remote_mainDoc, "blip", "embed", a.NamespaceName, remote_Id, new_Id );
                // Replace all instances of remote_Id in the local document with local_Id (for shapes)
                this.ReplaceAllRemoteID( remote_mainDoc, "imagedata", "id", v.NamespaceName, remote_Id, new_Id );
              }
              else
              {
                // Using same image in header/footers for document1 and document2.
                this.CreateRelationshipForImage( part.Uri, remote_Id, remote_mainDoc, packagePartToSaveTo );
              }

              break;
            }
          }
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
            HelperFunctions.CopyStream( s_read, s_write );
          }
        }

        this.CreateRelationshipForImage( new Uri( new_uri, UriKind.Relative ), remote_Id, remote_mainDoc, packagePartToSaveTo );
      }
    }

    private void CreateRelationshipForImage( Uri newUri, string remote_Id, XDocument remote_mainDoc, PackagePart packagePartToSaveTo )
    {
      var pr = packagePartToSaveTo.CreateRelationship( newUri, TargetMode.Internal, Document.RelationshipImage );
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

    private void CreateRelationshipForChart( Uri newUri, string remote_Id, XDocument remote_mainDoc )
    {
      var relID = this.GetNextFreeRelationshipID();
      var pr = this.PackagePart.CreateRelationship( newUri, TargetMode.Internal, Document.RelationshipChart, relID );

      // Replace all instances of remote_Id in the local document with local_Id
      this.ReplaceAllRemoteID( remote_mainDoc, "chart", "id", c.NamespaceName, remote_Id, pr.Id );
    }

    private void MergeCharts( PackagePart remote_pp, Document remote_document, XDocument remote_mainDoc )
    {
      var remote_rel = remote_document.PackagePart.GetRelationships().Where( r => r.TargetUri.OriginalString.Equals( remote_pp.Uri.OriginalString.Replace( "/word/", "" ) ) ).FirstOrDefault();

      if( remote_rel != null )
      {
        var remote_Id = remote_rel.Id;

        var new_uri = remote_pp.Uri.OriginalString;
        new_uri = new_uri.Remove( new_uri.LastIndexOf( "/" ) );
        new_uri += "/" + Guid.NewGuid();
        if( !new_uri.StartsWith( "/" ) )
        {
          new_uri = "/" + new_uri;
        }

        var new_pp = _package.CreatePart( new Uri( new_uri, UriKind.Relative ), remote_pp.ContentType, CompressionOption.Normal );

        using( Stream s_read = remote_pp.GetStream() )
        {
          using( Stream s_write = new PackagePartStream( new_pp.GetStream( FileMode.Create ) ) )
          {
            HelperFunctions.CopyStream( s_read, s_write );
          }
        }

        this.CreateRelationshipForChart( new Uri( new_uri, UriKind.Relative ), remote_Id, remote_mainDoc );

        XDocument chartDoc = null;
        using( var stream = new_pp.GetStream( FileMode.Open, FileAccess.Read ) )
        {
          using( var xr = XmlReader.Create( stream ) )
          {
            chartDoc = XDocument.Load( xr );
          }
        }
        var chartRelationships = remote_pp.GetRelationships();
        foreach( var rel in chartRelationships )
        {
          var relationshipCreated = new_pp.CreateRelationship( rel.TargetUri, rel.TargetMode, rel.RelationshipType );

          this.ReplaceAllRemoteID( chartDoc, "externalData", "id", c.NamespaceName, rel.Id, relationshipCreated.Id );
        }

        using( var stream = new PackagePartStream( new_pp.GetStream( FileMode.Open, FileAccess.Write ) ) )
        {
          using( XmlWriter xw = XmlWriter.Create( stream ) )
          {
            chartDoc.WriteTo( xw );
          }
        }
      }
    }


    private void ReplaceAllRemoteID( XDocument remote_mainDoc, string localName, string localNameAttribute, string namespaceName, string remote_Id, string new_Id )
    {
      // Replace all instances of remote_Id in the local document with local_Id
      var elems = remote_mainDoc.Descendants( XName.Get( localName, namespaceName ) );
      foreach( var elem in elems )
      {
        var attribute = elem.Attribute( XName.Get( localNameAttribute, Document.r.NamespaceName ) );
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

    private void merge_endnotes( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, Document remote, XDocument remote_endnotes )
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
        if( id != null && HelperFunctions.TryParseInt( id.Value, out i ) )
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

    private void merge_footnotes( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, Document remote, XDocument remote_footnotes )
    {
      IEnumerable<int> ids =
      (
          from d in _footnotes.Root.Descendants()
          where d.Name.LocalName == "footnote"
          select int.Parse( d.Attribute( XName.Get( "id", Document.w.NamespaceName ) ).Value )
      );

      int max_id = ids.Max() + 1;
      var footnoteReferences = remote_mainDoc.Descendants( XName.Get( "footnoteReference", Document.w.NamespaceName ) );

      foreach( var footnote in remote_footnotes.Root.Elements().OrderBy( fr => fr.Attribute( XName.Get( "id", Document.r.NamespaceName ) ) ).Reverse() )
      {
        XAttribute id = footnote.Attribute( XName.Get( "id", Document.w.NamespaceName ) );
        int i;
        if( id != null && HelperFunctions.TryParseInt( id.Value, out i ) )
        {
          if( i > 0 )
          {
            foreach( var footnoteRef in footnoteReferences )
            {
              XAttribute a = footnoteRef.Attribute( XName.Get( "id", Document.w.NamespaceName ) );
              if( a != null && int.Parse( a.Value ).Equals( i ) )
              {
                a.SetValue( max_id );
              }
            }

            // We care about copying this footnote.
            footnote.SetAttributeValue( XName.Get( "id", Document.w.NamespaceName ), max_id );
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

    private void merge_numbering( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, Document remote )
    {
      // Add each remote numbering to this document.
      var abstractNumElement = _numbering.Root.Elements( XName.Get( "abstractNum", w.NamespaceName ) );
      var remote_abstractNums = remote._numbering.Root.Elements( XName.Get( "abstractNum", w.NamespaceName ) );

      int abstractNumGuid = -1;
      foreach( var an in abstractNumElement )
      {
        var a = an.Attribute( XName.Get( "abstractNumId", w.NamespaceName ) );
        if( a != null )
        {
          int i;
          if( HelperFunctions.TryParseInt( a.Value, out i ) )
          {
            if( i > abstractNumGuid )
            {
              abstractNumGuid = i;
            }
          }
        }
      }
      abstractNumGuid++;

      var remote_nums = remote._numbering.Root.Elements( XName.Get( "num", w.NamespaceName ) );
      var numElement = _numbering.Root.Elements( XName.Get( "num", w.NamespaceName ) );

      int numGuid = 0;
      foreach( var an in numElement )
      {
        var a = an.Attribute( XName.Get( "numId", w.NamespaceName ) );
        if( a != null )
        {
          int i;
          if( HelperFunctions.TryParseInt( a.Value, out i ) )
          {
            if( i > numGuid )
            {
              numGuid = i;
            }
          }
        }
      }
      numGuid++;

      foreach( XElement remote_abstractNum in remote_abstractNums )
      {
        var abstractNumId = remote_abstractNum.Attribute( XName.Get( "abstractNumId", w.NamespaceName ) );
        if( abstractNumId != null )
        {
          var abstractNumIdValue = abstractNumId.Value;
          abstractNumId.SetValue( abstractNumGuid );

          foreach( XElement remote_num in remote_nums )
          {
            //abstractNumId of this remote_num
            var e = remote_num.Element( XName.Get( "abstractNumId", w.NamespaceName ) );
            var a2 = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( a2 != null && a2.Value.Equals( abstractNumIdValue ) )
            {
              a2.SetValue( abstractNumGuid );

              var currentNumId = remote_num.Attribute( XName.Get( "numId", w.NamespaceName ) ).Value;

              // in document
              var numIds = remote_mainDoc.Descendants( XName.Get( "numId", w.NamespaceName ) );
              foreach( var numId in numIds )
              {
                var attr = numId.Attribute( XName.Get( "val", w.NamespaceName ) );
                if( attr != null && attr.Value.Equals( currentNumId ) )
                {
                  attr.SetValue( numGuid );
                }
              }
              // in styles
              numIds = remote._styles.Descendants( XName.Get( "numId", w.NamespaceName ) );
              foreach( var numId in numIds )
              {
                var attr = numId.Attribute( XName.Get( "val", w.NamespaceName ) );
                if( attr != null && attr.Value.Equals( currentNumId ) )
                {
                  attr.SetValue( numGuid );
                }
              }

              remote_num.SetAttributeValue( XName.Get( "numId", w.NamespaceName ), numGuid );
              numGuid++;
            }
          }
        }

        abstractNumGuid++;
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

    private void merge_fonts( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, Document remote )
    {
      // Add each remote font to this document.
      IEnumerable<XElement> remote_fonts = remote._fontTable.Root.Elements( XName.Get( "font", Document.w.NamespaceName ) );
      IEnumerable<XElement> local_fonts = _fontTable.Root.Elements( XName.Get( "font", Document.w.NamespaceName ) );

      foreach( XElement remote_font in remote_fonts )
      {
        bool flag_addFont = true;
        foreach( XElement local_font in local_fonts )
        {
          if( local_font.Attribute( XName.Get( "name", Document.w.NamespaceName ) ).Value == remote_font.Attribute( XName.Get( "name", Document.w.NamespaceName ) ).Value )
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

    private bool IsDefaultParagraphStyle( XElement style )
    {
      if( style == null )
        return false;

      return ( ( style.Attribute( XName.Get( "default", w.NamespaceName ) ) != null )
         && ( style.Attribute( XName.Get( "default", w.NamespaceName ) ).Value == "1" )
         && ( style.Attribute( XName.Get( "type", w.NamespaceName ) ) != null )
         && ( style.Attribute( XName.Get( "type", w.NamespaceName ) ).Value == "paragraph" ) );
    }

    private void MergeDefaultParagraphStyles( Document remote, XElement remote_style )
    {
      if( ( remote == null ) || ( remote_style == null ) )
        return;

      var defaults = remote.GetDocDefaults();
      if( defaults != null )
      {
        var remote_rPr_default = defaults.Element( XName.Get( "rPrDefault", w.NamespaceName ) );
        if( remote_rPr_default != null )
        {
          // Update remote document default paragraph with default remote document values.
          // This way, the remote default document values won't be necessary anymore.
          // The final merged document default paragraph will be used by initial document only (not by remote document).
          var remoteDefault_rPr = remote_rPr_default.Element( XName.Get( "rPr", w.NamespaceName ) );
          if( remoteDefault_rPr != null )
          {
            // Update the rPr
            var remoteDefaultStyleParagraph_rPr = remote_style.Element( XName.Get( "rPr", w.NamespaceName ) );
            if( remoteDefaultStyleParagraph_rPr == null )
            {
              // No rPr in defaultParagraph of document. Copy the document's default rPr into the defaultParagraph of document.
              remote_style.Add( remoteDefault_rPr );
            }
            else
            {
              // rPr exists in defaultParagraph of document. Copy only the missing rPr element from document's default into the defaultParagraph of document.
              foreach( var rPr in remoteDefault_rPr.Elements() )
              {
                var rPrElement = remoteDefaultStyleParagraph_rPr.Element( rPr.Name );
                if( rPrElement == null )
                {
                  // Add the missing rPr from defaultDocument to defaultParagraph.
                  remoteDefaultStyleParagraph_rPr.Add( rPr );
                }
              }
            }
          }
        }

        var remote_pPr_default = defaults.Element( XName.Get( "pPrDefault", w.NamespaceName ) );
        if( remote_pPr_default != null )
        {
          // Update document default paragraph with default document values.
          var remoteDefault_pPr = remote_pPr_default.Element( XName.Get( "pPr", w.NamespaceName ) );
          if( remoteDefault_pPr != null )
          {
            // Update the pPr
            var remoteDefaultStyleParagraph_pPr = remote_style.Element( XName.Get( "pPr", w.NamespaceName ) );
            if( remoteDefaultStyleParagraph_pPr == null )
            {
              // No pPr in defaultParagraph of document. Copy the document's default pPr into the defaultParagraph of document.
              remote_style.Add( remoteDefault_pPr );
            }
            else
            {
              // pPr exists in defaultParagraph of document. Copy only the missing pPr element from document's default into the defaultParagraph of document.
              foreach( var pPr in remoteDefault_pPr.Elements() )
              {
                var pPrElement = remoteDefaultStyleParagraph_pPr.Element( pPr.Name );
                if( pPrElement == null )
                {
                  // Add the missing pPr from defaultDocument to defaultParagraph.
                  remoteDefaultStyleParagraph_pPr.Add( pPr );
                }
              }
            }
          }
        }
      }
    }

    private void merge_styles( PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, Document remote, XDocument remote_footnotes, XDocument remote_endnotes, MergingMode mergingMode )
    {
      var local_styles = new Dictionary<string, XElement>();

      foreach( var local_style in _styles.Root.Elements( XName.Get( "style", w.NamespaceName ) ) )
      {
        var styleId = local_style.Attribute( XName.Get( "styleId", w.NamespaceName ) );
        var styleIdValue = styleId.Value;

        local_styles.Add( styleIdValue, local_style );
      }

      // Add each remote style to this document.
      var remote_styles = remote._styles.Root.Elements( XName.Get( "style", w.NamespaceName ) );

      foreach( var remote_style in remote_styles )
      {
        String guuid;
        var styleId = remote_style.Attribute( XName.Get( "styleId", w.NamespaceName ) );
        var styleIdValue = styleId.Value;

        // Check to see if the local document already contains the remote styleId.
        if( local_styles.ContainsKey( styleIdValue ) )
        {
          switch( mergingMode )
          {
            case MergingMode.Local:
              {
                // Keep the local document style : nothing to do.
                continue;
              }
            case MergingMode.Remote:
              {
                // Replace the local document style with the remote document style.
                var sameLocalStyle = local_styles[ styleIdValue ];
                if( sameLocalStyle != null )
                {
                  sameLocalStyle.AddAfterSelf( remote_style );
                  sameLocalStyle.Remove();
                  local_styles[ styleIdValue ] = remote_style;
                }
                continue;
              }
            case MergingMode.Both:
              {
                // If local style and remote style are equals, nothing to do.
                var sameLocalStyle = local_styles[ styleIdValue ];
                if( sameLocalStyle.ToString() == remote_style.ToString() )
                {
                  continue;
                }
              }
              break;
            default:
              {
                Debug.Assert( false, "Unknown style selection type." );
                break;
              }
          }
        }

        // Create a new style from the remote style and add it to the local document style.
        guuid = Guid.NewGuid().ToString();

        var found = false;

        foreach( XElement e in remote_mainDoc.Root.Descendants( XName.Get( "pStyle", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
          {
            e_styleId.SetValue( guuid );
            found = true;
          }
        }

        foreach( XElement e in remote_mainDoc.Root.Descendants( XName.Get( "rStyle", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
          {
            e_styleId.SetValue( guuid );
            found = true;
          }
        }

        foreach( XElement e in remote_mainDoc.Root.Descendants( XName.Get( "tblStyle", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
          {
            e_styleId.SetValue( guuid );
            found = true;
          }

          var tablePrs = (
            from s in remote._styles.Element( Document.w + "styles" ).Elements( Document.w + "style" )
            let type = s.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
            where ( ( type != null ) && ( ( type.Value == "table" ) ) )
            select s.Element( XName.Get( "tblPr", Document.w.NamespaceName ) )
            );
          foreach( var tblPr in tablePrs )
          {
            var tblInd = tblPr.Element( XName.Get( "tblInd", Document.w.NamespaceName ) );
            if( tblInd != null )
            {
              var w = tblInd.Attribute( XName.Get( "w", Document.w.NamespaceName ) );
              if( ( w != null ) && ( w.Value == "0" ) )
              {
                tblInd.Remove();
              }
            }
          }
        }

        if( this.IsDefaultParagraphStyle( remote_style ) )
        {
          // Update remote document default paragraph with default remote document values.
          // This way, the remote default document values won't be necessary anymore.
          // The final merged document default paragraph will be used by initial document only (not by remote document).
          this.MergeDefaultParagraphStyles( remote, remote_style );

          remote_style.Attribute( XName.Get( "default", w.NamespaceName ) ).Remove();

          foreach( var p in remote_mainDoc.Root.Descendants( XName.Get( "p", w.NamespaceName ) ) )
          {
            if( ( p.Parent != null ) && ( p.Parent.Name != null ) && ( p.Parent.Name.LocalName == "tc" ) )
              continue;

            var pPr = p.Element( XName.Get( "pPr", w.NamespaceName ) );
            if( pPr == null )
            {
              p.Add( new XElement( XName.Get( "pPr", w.NamespaceName ) ) );
              pPr = p.Element( XName.Get( "pPr", w.NamespaceName ) );
            }

            var pStyle = pPr.Element( XName.Get( "pStyle", w.NamespaceName ) );
            if( pStyle == null )
            {
              pPr.Add( new XElement( XName.Get( "pStyle", w.NamespaceName ) ) );
              pStyle = pPr.Element( XName.Get( "pStyle", w.NamespaceName ) );
              pStyle.SetAttributeValue( XName.Get( "val", w.NamespaceName ), guuid );
            }
            else
            {
              var e_styleId = pStyle.Attribute( XName.Get( "val", w.NamespaceName ) );
              if( ( e_styleId != null ) && e_styleId.Value.Equals( styleId.Value ) )
              {
                e_styleId.SetValue( guuid );
              }
            }
          }
        }

        if( remote_endnotes != null )
        {
          foreach( XElement e in remote_endnotes.Root.Descendants( XName.Get( "rStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
            {
              e_styleId.SetValue( guuid );
              found = true;
            }
          }

          foreach( XElement e in remote_endnotes.Root.Descendants( XName.Get( "pStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
            {
              e_styleId.SetValue( guuid );
              found = true;
            }
          }
        }

        if( remote_footnotes != null )
        {
          foreach( XElement e in remote_footnotes.Root.Descendants( XName.Get( "rStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
            {
              e_styleId.SetValue( guuid );
              found = true;
            }
          }

          foreach( XElement e in remote_footnotes.Root.Descendants( XName.Get( "pStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
            {
              e_styleId.SetValue( guuid );
              found = true;
            }
          }
        }

        // Modify the styles in _numbering remote and local, to make sure the order of merge (style vs numbering) won't affect final local styles.
        if( remote._numbering != null )
        {
          foreach( XElement e in remote._numbering.Root.Descendants( XName.Get( "pStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
            {
              e_styleId.SetValue( guuid );
              found = true;
            }
          }
        }
        if( _numbering != null )
        {
          foreach( XElement e in _numbering.Root.Descendants( XName.Get( "pStyle", w.NamespaceName ) ) )
          {
            var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
            if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
            {
              e_styleId.SetValue( guuid );
              found = true;
            }
          }
        }

        foreach( var e in remote._styles.Descendants( XName.Get( "basedOn", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
          {
            e_styleId.SetValue( guuid );
            found = true;
          }
        }

        foreach( var e in remote._styles.Descendants( XName.Get( "next", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
          {
            e_styleId.SetValue( guuid );
            found = true;
          }
        }

        foreach( var e in remote._styles.Descendants( XName.Get( "link", w.NamespaceName ) ) )
        {
          var e_styleId = e.Attribute( XName.Get( "val", w.NamespaceName ) );
          if( ( e_styleId != null ) && e_styleId.Value.Equals( styleIdValue ) )
          {
            e_styleId.SetValue( guuid );
            found = true;
          }
        }

        // unsused style. do not add it. Too many style may cause a corrupted resulting document.
        if( !found )
          continue;

        // Set the styleId and name in the remote_style to this new Guid
        remote_style.SetAttributeValue( XName.Get( "styleId", w.NamespaceName ), guuid );

        if( local_styles.ContainsKey( styleIdValue ) )
        {
          var name = remote_style.Element( XName.Get( "name", w.NamespaceName ) );
          if( name != null )
          {
            name.SetAttributeValue( XName.Get( "val", w.NamespaceName ), guuid );
          }
        }

        _styles.Root.Add( remote_style );
      }
    }































    protected void clonePackageRelationship( Document remote_document, PackagePart pp, XDocument remote_mainDoc )
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
          HelperFunctions.CopyStream( s_read, s_write );
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

    private static void PopulateDocument( Document document, Package package )
    {
      document.Xml = document._mainDoc.Root.Element( XName.Get( "body", Document.w.NamespaceName ) );
      if( document.Xml == null )
        throw new InvalidDataException( "Can't find body of document's xml. Make sure document has a body from namespace w:http://schemas.openxmlformats.org/wordprocessingml/2006/main" );

      document._settingsPart = HelperFunctions.CreateOrGetSettingsPart( package );

      try
      {
        var rel = document.PackagePart.GetRelationships();
      }
      catch( UriFormatException )
      {
        Document.UpdateRelationshipsUri( package );
      }

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
            {
              document._styles = XDocument.Load( tr );
              var docDefaults = document._styles.Root.Element( XName.Get( "docDefaults", Document.w.NamespaceName ) );
              if( docDefaults != null )
              {
                var pPrDefault = docDefaults.Element( XName.Get( "pPrDefault", Document.w.NamespaceName ) );
                if( pPrDefault != null )
                {
                  var pPr = pPrDefault.Element( XName.Get( "pPr", Document.w.NamespaceName ) );
                  if( pPr != null )
                  {
                    Paragraph.SetDefaultValues( pPr );
                  }
                }
              }
            }
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

      document._cachedSections = document.GetSections();
    }

    // This method will parse the Hyperlinks of the document and replace the mal-formed ones with "http://broken-link/" 
    // in order for document.PackagePart.GetRelationships() to not throw exceptions.
    private static void UpdateRelationshipsUri( Package package )
    {
      if( package == null )
        return;

      if( package.PartExists( new Uri( "/word/_rels/document.xml.rels", UriKind.Relative ) ) )
      {
        XDocument docRelationShipDocument;
        bool needUpdate = false;
        var docRelationships = package.GetPart( new Uri( "/word/_rels/document.xml.rels", UriKind.Relative ) );
        using( var tr = new StreamReader( docRelationships.GetStream( FileMode.Open, FileAccess.Read ) ) )
        {
          docRelationShipDocument = XDocument.Load( tr, LoadOptions.PreserveWhitespace );

          var urisToValidate = docRelationShipDocument
                                  .Descendants( XName.Get( "Relationship", rel.NamespaceName ) )
                                  .Where( relation => ( relation.Attribute( "TargetMode" ) != null ) && ( (string)relation.Attribute( "TargetMode" ) == "External" ) );

          foreach( var relation in urisToValidate )
          {
            var target = (string)relation.Attribute( "Target" );
            if( !string.IsNullOrEmpty( target ) )
            {
              try
              {
                var uri = new Uri( target );
              }
              catch( UriFormatException )
              {
                var newUri = new Uri( "http://broken-link/" );
                relation.Attribute( "Target" ).Value = newUri.ToString();
                needUpdate = true;
              }
            }
          }
        }

        if( needUpdate )
        {
          using( var tw = new StreamWriter( new PackagePartStream( docRelationships.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            docRelationShipDocument.Save( tw, SaveOptions.None );
          }
        }
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
      if( HelperFunctions.TryParseInt( newId, out result ) )
        return ( "rId" + ( result + 1 ) );

      var guid = String.Empty;
      do
      {
        guid = Guid.NewGuid().ToString();
      } while( Char.IsDigit( guid[ 0 ] ) );
      return guid;
    }







    private Hyperlink AddHyperlinkCore( string text, Uri uri, string anchor, Hyperlink baseHyperlink, Formatting formatting )
    {
      XElement xElement = null;




      {
        xElement = new XElement
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
      }

      var h = new Hyperlink( this, this.PackagePart, xElement );
      h.text = text;
      if( uri != null )
      {
        h.uri = uri;
      }

      this.AddHyperlinkStyleIfNotPresent();

      return h;
    }

    private Section InsertSection( bool trackChanges, bool isPageBreak )
    {
      if( isPageBreak )
      {
        return this.Sections.Last().InsertSectionPageBreak( trackChanges );
      }
      else
      {
        return this.Sections.Last().InsertSection( trackChanges );
      }
    }

    private void ManageSectionBreak( bool append, bool useSectionBreak, XElement remote_body, XElement local_body )
    {
      if( useSectionBreak )
      {
        // append : Move local body section to the last paragraph of the body.
        // insert : Move Remote Body section to the last paragraph of the body.
        this.MoveSectionIntoLastParagraph( append ? local_body : remote_body );
      }
      else
      {
        if( append )
        {
          // The last section of local will become the last section of remote(will be the last section of the resulting document).
          this.ReplaceLastSection( local_body, remote_body );
        }
        else
        {
          // The last section of remote is removed. The last section of local will be the last section of the resulting document.
          this.RemoveLastSection( remote_body );
        }
      }
    }

    private void MoveSectionIntoLastParagraph( XElement body )
    {
      Debug.Assert( body != null, "body shouldn't be null." );

      var body_sectPr = body.Elements( XName.Get( "sectPr", Document.w.NamespaceName ) ).LastOrDefault();
      if( body_sectPr != null )
      {
        var bodyLastParagraph = body.Elements( XName.Get( "p", Document.w.NamespaceName ) ).LastOrDefault();
        if( bodyLastParagraph == null )
        {
          body.Add( new XElement( XName.Get( "p", Document.w.NamespaceName ) ) );
          bodyLastParagraph = body.Elements( XName.Get( "p", Document.w.NamespaceName ) ).LastOrDefault();
        }
        var bodyLastParagraphProperties = bodyLastParagraph.Element( XName.Get( "pPr", Document.w.NamespaceName ) );
        if( bodyLastParagraphProperties == null )
        {
          bodyLastParagraph.Add( new XElement( XName.Get( "pPr", Document.w.NamespaceName ) ) );
          bodyLastParagraphProperties = bodyLastParagraph.Element( XName.Get( "pPr", Document.w.NamespaceName ) );
        }
        var bodyLastParagraph_sectPr = bodyLastParagraphProperties.Element( XName.Get( "sectPr", Document.w.NamespaceName ) );
        if( bodyLastParagraph_sectPr == null )
        {
          bodyLastParagraphProperties.Add( body_sectPr );
        }
        body_sectPr.Remove();
      }
    }

    private void RemoveLastSection( XElement body )
    {
      Debug.Assert( body != null, "body shouldn't be null." );

      var body_sectPr = body.Elements( XName.Get( "sectPr", Document.w.NamespaceName ) ).LastOrDefault();
      if( body_sectPr != null )
      {
        body_sectPr.Remove();
      }
    }

    private void ReplaceLastSection( XElement first, XElement second )
    {
      Debug.Assert( first != null, "first shouldn't be null." );
      Debug.Assert( second != null, "second shouldn't be null." );

      var first_sectPr = first.Elements( XName.Get( "sectPr", Document.w.NamespaceName ) ).LastOrDefault();
      if( first_sectPr != null )
      {
        var second_sectPr = second.Elements( XName.Get( "sectPr", Document.w.NamespaceName ) ).LastOrDefault();
        if( second_sectPr != null )
        {
          second_sectPr.ReplaceWith( first_sectPr );
          first_sectPr.Remove();
        }
      }
    }

    private void InsertChart( Chart chart, Paragraph paragraph, float width = 432f, float height = 252f )
    {
      var p = this.CreateChartElement( chart, paragraph, width, height );
      p.Xml.Add( chart.Xml );

      // Extract the attribute id from the Chart Xml.
      var docPr =
      (
          from e in p.Xml.Elements().Last().Descendants()
          where e.Name.LocalName.Equals( "docPr" )
          select e
      ).Single();

      // Set its value to a unique id.
      docPr.SetAttributeValue( "id", this.Document.GetNextFreeDocPrId().ToString() );
    }

    private void UpdateNoteXmlFromContent( XElement noteXml, Paragraph noteParagraph, object footnoteContent, Formatting formatting )
    {
      Debug.Assert( noteXml != null, "noteXml should not be null." );
      Debug.Assert( noteParagraph != null, "noteParagraph should not be null." );

      if( footnoteContent is string )
      {
        noteParagraph.InsertText( footnoteContent as string, false, formatting );
      }
      else if( footnoteContent is Hyperlink )
      {
        noteParagraph.InsertHyperlink( footnoteContent as Hyperlink, noteParagraph.Text.Length );
      }
      else if( footnoteContent is Picture )
      {
        noteParagraph.InsertPicture( footnoteContent as Picture, noteParagraph.Text.Length );
      }
      else if( !( footnoteContent is Table ) )
      {
        throw new ArgumentException( "Unknown object received for footnote/endnote. Valid objects are string, Picture, Hyperlink or Table. If other configuration is needed, try using the Note.Paragraphs property and configure the Footnote/Endnote paragraphs." );
      }

      noteXml.Add( noteParagraph.Xml );

      // Tables need to be added after the newFootnoteParagraph is included inside the newFootnoteXml. This is because table is located after paragraph.
      if( footnoteContent is Table )
      {
        noteParagraph.InsertTableAfterSelf( footnoteContent as Table );
      }
    }

    private int GetMaxId( XElement xElement )
    {
      if( xElement == null )
        return -1;

      var lastId = xElement.Attribute( XName.Get( "id", Document.w.NamespaceName ) );
      if( lastId == null )
        return -1;

      return Int32.Parse( lastId.Value ) + 1;
    }

    private HashSet<string> GetExistingDocPrIds( XElement xElement )
    {
      var result = new HashSet<string>();

      if( xElement == null )
        return result;

      var xNameBookmarkStart = XName.Get( "bookmarkStart", Document.w.NamespaceName );
      var xNameDocPr = XName.Get( "docPr", Document.wp.NamespaceName );

      foreach( var id in xElement.Descendants() )
      {
        if( ( id.Name != xNameBookmarkStart ) && ( id.Name != xNameDocPr ) )
          continue;

        var idAtt = id.Attributes().FirstOrDefault( x => x.Name.LocalName == "id" );
        if( idAtt != null )
        {
          result.Add( idAtt.Value );
        }
      }

      return result;
    }






    #endregion

    #region Constructors

    internal Document( Document document, XElement xml )
      : base( document, xml )
    {
      Paragraph.ResetDefaultValues();
    }

    #endregion

    #region IDisposable Members

    public void Dispose()
    {
      _package.Close();
    }

    #endregion
  }

  public enum HashAlgorithm
  {
    None,
    Md2,
    Md5,
    Sha1,
    Sha256,
    Sha384,
    Sha512
  }


  public enum HashAlgorithmName
  {
    NotSupported,
    SHA1,
    SHA1Managed,
    SHA256Managed,
    SHA384Managed,
    SHA512Managed
  }

  internal class Encryption
  {
    #region Public fields

    public int Iterations;
    public byte[] KeyValues;
    public byte[] SaltArray;

    #endregion

    #region Public Methods
    public Encryption Encrypt( string password, HashAlgorithmName hashAlgorithm )
    {
      return EncryptOrDecrypt( password, "", hashAlgorithm );
    }

    public string Decrypt( string password, string salt, HashAlgorithmName hashAlgorithm )
    {
      var decrypt = EncryptOrDecrypt( password, salt, hashAlgorithm );
      return Convert.ToBase64String( decrypt.KeyValues );
    }

    #endregion

    #region private methods

    private static System.Security.Cryptography.HashAlgorithm GetHashAlgorithm( HashAlgorithmName algorithm )
    {
      switch( algorithm )
      {
        case HashAlgorithmName.SHA1:
          return System.Security.Cryptography.SHA1.Create();
        case HashAlgorithmName.SHA1Managed:
          return System.Security.Cryptography.SHA1Managed.Create();
        case HashAlgorithmName.SHA256Managed:
          return System.Security.Cryptography.SHA256Managed.Create();
        case HashAlgorithmName.SHA384Managed:
          return System.Security.Cryptography.SHA384Managed.Create();
        case HashAlgorithmName.SHA512Managed:
          return System.Security.Cryptography.SHA512Managed.Create();
        default:
          throw new NotSupportedException( "Unsupported password encryption algorithm" );
      }
    }

    private Encryption EncryptOrDecrypt( string password, string salt = "", HashAlgorithmName hashAlgorithm = HashAlgorithmName.NotSupported )
    {
      // Intellectual Property information :
      //
      // The following code handles password protection of Word documents (Open Specifications) 
      // and is an implementation of algorithm(s) described in Office Document Cryptography Structure 
      // here: https://msdn.microsoft.com/en-us/library/cc313071.aspx.
      //
      // The code’s use is covered under Microsoft’s Open Specification Promise 
      // described here: https://msdn.microsoft.com/en-US/openspecifications/dn646765

      int maxPasswordLength = 15;
      var saltArray = !string.IsNullOrEmpty( salt ) ? Convert.FromBase64String( salt ) : new byte[ 16 ];
      var keyValues = new byte[ 14 ];

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

      if( saltArray.All( item => item == 0 ) )
      {
        // Generate the salt
        var random = new RNGCryptoServiceProvider();
        random.GetNonZeroBytes( saltArray );
      }

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
          keyValues[ i ] = Convert.ToByte( ( (uint)( intCombinedkey & ( 0x000000FF << ( i * 8 ) ) ) ) >> ( i * 8 ) );
        }
      }

      var sb = new StringBuilder();
      for( int intTemp = 0; intTemp < 4; intTemp++ )
      {
        // Fix for Github issue #478 - https://github.com/xceedsoftware/DocX/issues/478
        // Convert to hexadecimal string with at least 2 digits, padding with leading zeros if necessary
        // Otherwise computed hash value will fail to match computed keyValues 
        var str = keyValues[ intTemp ].ToString( "x2" );
        sb.Append( str );
      }

      keyValues = Encoding.Unicode.GetBytes( sb.ToString().ToUpper() );
      keyValues = MergeArrays( keyValues, saltArray );

      int iterations = 100000;

      var hash = Encryption.GetHashAlgorithm( hashAlgorithm );
      keyValues = hash.ComputeHash( keyValues );
      var iterator = new byte[ 4 ];
      for( int i = 0; i < iterations; i++ )
      {
        iterator[ 0 ] = Convert.ToByte( ( i & 0x000000FF ) >> 0 );
        iterator[ 1 ] = Convert.ToByte( ( i & 0x0000FF00 ) >> 8 );
        iterator[ 2 ] = Convert.ToByte( ( i & 0x00FF0000 ) >> 16 );
        iterator[ 3 ] = Convert.ToByte( ( i & 0xFF000000 ) >> 24 );

        keyValues = MergeArrays( iterator, keyValues );
        keyValues = hash.ComputeHash( keyValues );
      }

      Encryption encryption = new Encryption()
      {
        Iterations = iterations,
        KeyValues = keyValues,
        SaltArray = saltArray
      };

      return encryption;
    }

    private byte[] MergeArrays( byte[] array1, byte[] array2 )
    {
      byte[] result = new byte[ array1.Length + array2.Length ];
      Buffer.BlockCopy( array2, 0, result, 0, array2.Length );
      Buffer.BlockCopy( array1, 0, result, array2.Length, array1.Length );
      return result;
    }

    #endregion
  }
}
