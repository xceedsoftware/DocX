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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.Xml.Linq;
using System.IO;
using System.Reflection;
using System.IO.Compression;
using System.Globalization;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace Xceed.Document.NET
{
  internal enum ResourceType
  {
    DefaultStyle,
    NumberingBullet,
    NumberingDecimal,
    Numbering,
    Styles,
    Theme
  }

  internal static class HelperFunctions
  {
    public const string DOCUMENT_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    public const string TEMPLATE_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
    public const string SETTING_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    public const string MACRO_DOCUMENTTYPE = "application/vnd.ms-word.document.macroEnabled.main+xml";

    internal static readonly char[] RestrictedXmlCharacters = new char[]
    {
      '\x1','\x2','\x3','\x4','\x5','\x6','\x7','\x8','\xb','\xc','\xe','\xf',
      '\x10','\x11','\x12','\x13','\x14','\x15','\x16','\x17','\x18','\x19','\x1a','\x1b','\x1c','\x1e','\x1f',
      '\x7f','\x80','\x81','\x82','\x83','\x84','\x86','\x87','\x88','\x89','\x8a','\x8b','\x8c','\x8d','\x8e','\x8f',
      '\x90','\x91','\x92','\x93','\x94','\x95','\x96','\x97','\x98','\x99','\x9a','\x9b','\x9c','\x9d','\x9e','\x9f'
    };

    internal static void CreateRelsPackagePart( Document Document, Uri uri )
    {
      PackagePart pp = Document._package.CreatePart( uri, Document.ContentTypeApplicationRelationShipXml, CompressionOption.Maximum );
      using( TextWriter tw = new StreamWriter( new PackagePartStream( pp.GetStream() ) ) )
      {
        XDocument d = new XDocument
        (
            new XDeclaration( "1.0", "UTF-8", "yes" ),
            new XElement( XName.Get( "Relationships", Document.rel.NamespaceName ) )
        );
        var root = d.Root;
        d.Save( tw );
      }
    }

    internal static int GetSize( XElement Xml )
    {
      switch( Xml.Name.LocalName )
      {
        case "tab":
          return ( Xml.Parent.Name.LocalName != "tabs" ) ? 1 : 0;
        case "br":
          return ( HelperFunctions.IsLineBreak( Xml ) ) ? 1 : 0;
        case "t":
          goto case "delText";
        case "delText":
          return Xml.Value.Length;
        case "tr":
          goto case "br";
        case "tc":
          goto case "br";
        case "ptab":
          return 1;
        default:
          return 0;
      }
    }

    internal static string GetText( XElement e )
    {
      StringBuilder sb = new StringBuilder();
      GetTextRecursive( e, ref sb );
      return sb.ToString();
    }

    internal static void GetTextRecursive( XElement Xml, ref StringBuilder sb )
    {
      sb.Append( ToText( Xml ) );

      // Do not read text from Fallback or drawing(a new paragraph will take care of drawing).
      var fallbackValue = Xml.Name.Equals( XName.Get( "Fallback", Document.mc.NamespaceName ) );
      var drawingValue = Xml.Name.Equals( XName.Get( "drawing", Document.w.NamespaceName ) );

      if( Xml.HasElements && !fallbackValue && !drawingValue )
        foreach( XElement e in Xml.Elements() )
          GetTextRecursive( e, ref sb );
    }

    internal static List<FormattedText> GetFormattedText( XElement e )
    {
      List<FormattedText> alist = new List<FormattedText>();
      GetFormattedTextRecursive( e, ref alist );
      return alist;
    }

    internal static void GetFormattedTextRecursive( XElement Xml, ref List<FormattedText> alist )
    {
      FormattedText ft = ToFormattedText( Xml );
      FormattedText last = null;

      if( ft != null )
      {
        if( alist.Count() > 0 )
          last = alist.Last();

        if( last != null && last.CompareTo( ft ) == 0 )
        {
          // Update text of last entry.
          last.text += ft.text;
        }

        else
        {
          if( last != null )
            ft.index = last.index + last.text.Length;

          alist.Add( ft );
        }
      }

      if( Xml.HasElements )
        foreach( XElement e in Xml.Elements() )
          GetFormattedTextRecursive( e, ref alist );
    }

    internal static FormattedText ToFormattedText( XElement e )
    {
      // The text representation of e.
      String text = ToText( e );
      if( text == String.Empty )
        return null;

      // Do not read text from inner Fallback.
      var fallbackValue = e.AncestorsAndSelf().FirstOrDefault( x => x.Name.Equals( XName.Get( "Fallback", Document.mc.NamespaceName ) ) );
      if( fallbackValue != null )
        return null;

      // Do not read text from inner AlternateContent.
      //var alternateContentValue = e.AncestorsAndSelf().FirstOrDefault( x => x.Name.Equals( XName.Get( "AlternateContent", Document.mc.NamespaceName ) ) );
      //if( alternateContentValue != null )
      //  return null;

      // e is a w:t element, it must exist inside a w:r element or a w:tabs, lets climb until we find it.
      while( ( e != null ) && !e.Name.Equals( XName.Get( "r", Document.w.NamespaceName ) ) && !e.Name.Equals( XName.Get( "tabs", Document.w.NamespaceName ) ) )
        e = e.Parent;

      FormattedText ft = new FormattedText();
      ft.text = text;
      ft.index = 0;
      ft.formatting = null;

      if( e != null )
      {
        // e is a w:r element, lets find the rPr element.
        XElement rPr = e.Element( XName.Get( "rPr", Document.w.NamespaceName ) );

        // Return text with formatting.
        if( rPr != null )
        {
          ft.formatting = Formatting.Parse( rPr );
        }
      }

      return ft;
    }

    internal static bool IsLineBreak( XElement xml )
    {
      if( xml == null )
        return false;

      if( xml.Name.LocalName != "br" )
        return false;

      var breakNodeValue = xml.HasAttributes && ( xml.Attribute( Document.w + "type" ) != null ) ? xml.Attribute( Document.w + "type" ).Value : null;
      return ( ( breakNodeValue == null ) || ( breakNodeValue == "textWrapping" ) );
    }

    internal static string ToText( XElement e )
    {
      switch( e.Name.LocalName )
      {
        case "tab":
          // Do not add "\t" for TabStopPositions defined in "tabs".
          return ( ( e.Parent != null ) && e.Parent.Name.Equals( XName.Get( "tabs", Document.w.NamespaceName ) ) ) ? "" : "\t";
        // absolute tab
        case "ptab":
          return "\v";
        case "br":
          {
            // Manage only line Breaks.
            if( HelperFunctions.IsLineBreak( e ) )
              return "\n";

            return "";
          }
        case "t":
          goto case "delText";
        case "delText":
          {


            return e.Value;
          }
        case "tr":
          //goto case "br";
          return "\n";
        case "tc":
          goto case "tab";
        default:
          return "";
      }
    }

    internal static XElement CloneElement( XElement element )
    {
      return new XElement
      (
          element.Name,
          element.Attributes(),
          element.Nodes().Select
          (
              n =>
              {
                XElement e = n as XElement;
                if( e != null )
                  return CloneElement( e );
                return n;
              }
          )
      );
    }

    internal static PackagePart GetMainDocumentPart( Package package )
    {
      return package.GetParts().Single( p => p.ContentType.Equals( DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase ) ||
                                             p.ContentType.Equals( TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase ) ||
                                             p.ContentType.Equals( MACRO_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase ) );
    }

    internal static PackagePart CreateOrGetSettingsPart( Package package )
    {
      PackagePart settingsPart;

      var settingsUri = new Uri( "/word/settings.xml", UriKind.Relative );
      if( !package.PartExists( settingsUri ) )
      {
        settingsPart = package.CreatePart( settingsUri, HelperFunctions.SETTING_DOCUMENTTYPE, CompressionOption.Maximum );

        var mainDocPart = GetMainDocumentPart( package );

        mainDocPart.CreateRelationship( settingsUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" );

        var settings = XDocument.Parse
        ( @"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                <w:settings xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w10='urn:schemas-microsoft-com:office:word' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:sl='http://schemas.openxmlformats.org/schemaLibrary/2006/main'>
                  <w:zoom w:percent='100' />
                  <w:defaultTabStop w:val='720' />
                  <w:characterSpacingControl w:val='doNotCompress' />
                  <w:compat />
                  <w:rsids>
                    <w:rsidRoot w:val='00217F62' />
                    <w:rsid w:val='001915A3' />
                    <w:rsid w:val='00217F62' />
                    <w:rsid w:val='00A906D8' />
                    <w:rsid w:val='00AB5A74' />
                    <w:rsid w:val='00F071AE' />
                  </w:rsids>
                  <m:mathPr>
                    <m:mathFont m:val='Cambria Math' />
                    <m:brkBin m:val='before' />
                    <m:brkBinSub m:val='--' />
                    <m:smallFrac m:val='off' />
                    <m:dispDef />
                    <m:lMargin m:val='0' />
                    <m:rMargin m:val='0' />
                    <m:defJc m:val='centerGroup' />
                    <m:wrapIndent m:val='1440' />
                    <m:intLim m:val='subSup' />
                    <m:naryLim m:val='undOvr' />
                  </m:mathPr>
                  <w:displayBackgroundShape w:val='true' />
                  <w:themeFontLang w:val='en-IE' w:bidi='ar-SA' />
                  <w:clrSchemeMapping w:bg1='light1' w:t1='dark1' w:bg2='light2' w:t2='dark2' w:accent1='accent1' w:accent2='accent2' w:accent3='accent3' w:accent4='accent4' w:accent5='accent5' w:accent6='accent6' w:hyperlink='hyperlink' w:followedHyperlink='followedHyperlink' />
                  <w:shapeDefaults>
                    <o:shapedefaults v:ext='edit' spidmax='2050' />
                    <o:shapelayout v:ext='edit'>
                      <o:idmap v:ext='edit' data='1' />
                    </o:shapelayout>
                  </w:shapeDefaults>
                  <w:decimalSymbol w:val='.' />
                  <w:listSeparator w:val=',' />
                </w:settings>"
        );

        var themeFontLang = settings.Root.Element( XName.Get( "themeFontLang", Document.w.NamespaceName ) );
        themeFontLang.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), CultureInfo.CurrentCulture );

        // Save the settings document.
        using( TextWriter tw = new StreamWriter( new PackagePartStream( settingsPart.GetStream() ) ) )
        {
          settings.Save( tw );
        }
      }
      else
      {
        settingsPart = package.GetPart( settingsUri );
      }
      return settingsPart;
    }

    internal static void CreateCorePropertiesPart( Document document )
    {
      PackagePart corePropertiesPart = document._package.CreatePart( new Uri( "/docProps/core.xml", UriKind.Relative ), "application/vnd.openxmlformats-package.core-properties+xml", CompressionOption.Maximum );

      XDocument corePropDoc = XDocument.Parse( @"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
      <cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dcterms='http://purl.org/dc/terms/' xmlns:dcmitype='http://purl.org/dc/dcmitype/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'> 
         <dc:title></dc:title>
         <dc:subject></dc:subject>
         <dc:creator></dc:creator>
         <cp:keywords></cp:keywords>
         <dc:description></dc:description>
         <cp:lastModifiedBy></cp:lastModifiedBy>
         <cp:revision>1</cp:revision>
         <dcterms:created xsi:type='dcterms:W3CDTF'>" + DateTime.UtcNow.ToString( "s" ) + "Z" + @"</dcterms:created>
         <dcterms:modified xsi:type='dcterms:W3CDTF'>" + DateTime.UtcNow.ToString( "s" ) + "Z" + @"</dcterms:modified>
      </cp:coreProperties>" );

      using( TextWriter tw = new StreamWriter( new PackagePartStream( corePropertiesPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
        corePropDoc.Save( tw, SaveOptions.None );

      document._package.CreateRelationship( corePropertiesPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" );
    }

    internal static void CreateCustomPropertiesPart( Document document )
    {
      var customPropertiesPart = document._package.CreatePart( new Uri( "/docProps/custom.xml", UriKind.Relative ), "application/vnd.openxmlformats-officedocument.custom-properties+xml", CompressionOption.Maximum );

      var customPropDoc = new XDocument
      (
          new XDeclaration( "1.0", "UTF-8", "yes" ),
          new XElement
          (
              XName.Get( "Properties", Document.customPropertiesSchema.NamespaceName ),
              new XAttribute( XNamespace.Xmlns + "vt", Document.customVTypesSchema )
          )
      );

      using( TextWriter tw = new StreamWriter( new PackagePartStream( customPropertiesPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        customPropDoc.Save( tw, SaveOptions.None );
      }

      document._package.CreateRelationship( customPropertiesPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" );
    }

    internal static XDocument DecompressXMLResource( string manifest_resource_name )
    {
      // XDocument to load the compressed Xml resource into.
      XDocument document;

      // Get a reference to the executing assembly.
      Assembly assembly = Assembly.GetExecutingAssembly();

      // Open a Stream to the embedded resource.
      Stream stream = assembly.GetManifestResourceStream( manifest_resource_name );

      // Decompress the embedded resource.
      using( GZipStream zip = new GZipStream( stream, CompressionMode.Decompress ) )
      {
        // Load this decompressed embedded resource into an XDocument using a TextReader.
        using( TextReader sr = new StreamReader( zip ) )
        {
          document = XDocument.Load( sr );
        }
      }

      // Return the decompressed Xml as an XDocument.
      return document;
    }

    internal static string GetResources( ResourceType resType )
    {
      switch( resType )
      {
        case ResourceType.DefaultStyle:
          {
#if NETSTANDARD
            return "Xceed.Document.NETStandard.Resources.default_styles.xml.gz";
#else
            return "Xceed.Document.NET.Resources.default_styles.xml.gz";
#endif
          }
        case ResourceType.Numbering:
          {
#if NETSTANDARD
            return "Xceed.Document.NETStandard.Resources.numbering.xml.gz";
#else
            return "Xceed.Document.NET.Resources.numbering.xml.gz";
#endif
          }
        case ResourceType.NumberingBullet:
          {
#if NETSTANDARD
            return "Xceed.Document.NETStandard.Resources.numbering.default_bullet_abstract.xml.gz";
#else
            return "Xceed.Document.NET.Resources.numbering.default_bullet_abstract.xml.gz";
#endif
          }
        case ResourceType.NumberingDecimal:
          {
#if NETSTANDARD
            return "Xceed.Document.NETStandard.Resources.numbering.default_decimal_abstract.xml.gz";
#else
            return "Xceed.Document.NET.Resources.numbering.default_decimal_abstract.xml.gz";
#endif
          }
        case ResourceType.Styles:
          {
#if NETSTANDARD
            return "Xceed.Document.NETStandard.Resources.styles.xml.gz";
#else
            return "Xceed.Document.NET.Resources.styles.xml.gz";
#endif
          }
        case ResourceType.Theme:
          {
#if NETSTANDARD
            return "Xceed.Document.NETStandard.Resources.theme.xml.gz";
#else
            return "Xceed.Document.NET.Resources.theme.xml.gz";
#endif
          }
      }

      Debug.Assert( false, "Unkown resource type." );
      return null;
    }

    /// <summary>
    /// If this document does not contain a /word/styles.xml add the default one generated by Microsoft Word.
    /// </summary>
    /// <param name="package"></param>
    /// <returns></returns>
    internal static XDocument AddDefaultStylesXml( Package package )
    {
      XDocument stylesDoc;
      // Create the main document part for this package
      var word_styles = package.CreatePart( new Uri( "/word/styles.xml", UriKind.Relative ), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum );

      stylesDoc = HelperFunctions.DecompressXMLResource( HelperFunctions.GetResources( ResourceType.DefaultStyle ) );
      var lang = stylesDoc.Root.Element( XName.Get( "docDefaults", Document.w.NamespaceName ) ).Element( XName.Get( "rPrDefault", Document.w.NamespaceName ) ).Element( XName.Get( "rPr", Document.w.NamespaceName ) ).Element( XName.Get( "lang", Document.w.NamespaceName ) );
      lang.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), CultureInfo.CurrentCulture );

      // Save /word/styles.xml
      using( TextWriter tw = new StreamWriter( new PackagePartStream( word_styles.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        stylesDoc.Save( tw, SaveOptions.None );
      }

      var mainDocumentPart = GetMainDocumentPart( package );

      mainDocumentPart.CreateRelationship( word_styles.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" );
      return stylesDoc;
    }

    internal static XElement CreateEdit( EditType t, DateTime edit_time, object content )
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

      return
      (
          new XElement( Document.w + t.ToString(),
              new XAttribute( Document.w + "id", 0 ),
              new XAttribute( Document.w + "author", Environment.UserDomainName + "\\" + Environment.UserName ),
              new XAttribute( Document.w + "date", edit_time ),
          content )
      );
    }

    internal static XElement CreateTable( int rowCount, int columnCount )
    {
      if( ( rowCount <= 0 ) || ( columnCount <= 0 ) )
      {
        throw new ArgumentOutOfRangeException( "Row and Column count must be greater than 0." );
      }

      int[] columnWidths = new int[ columnCount ];
      for( int i = 0; i < columnCount; i++ )
      {
        columnWidths[ i ] = 2310;
      }
      return CreateTable( rowCount, columnWidths );
    }

    internal static XElement CreateTable( int rowCount, int[] columnWidths )
    {
      var newTable = new XElement( XName.Get( "tbl", Document.w.NamespaceName ),
                                   new XElement( XName.Get( "tblPr", Document.w.NamespaceName ),
                                                 new XElement( XName.Get( "tblStyle", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "TableGrid" ) ),
                                                 new XElement( XName.Get( "tblW", Document.w.NamespaceName ), new XAttribute( XName.Get( "w", Document.w.NamespaceName ), "5000" ), new XAttribute( XName.Get( "type", Document.w.NamespaceName ), "auto" ) ),
                                                 new XElement( XName.Get( "tblLook", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "04A0" ) ) ) );

      for( int i = 0; i < rowCount; i++ )
      {
        var row = new XElement( XName.Get( "tr", Document.w.NamespaceName ) );

        for( int j = 0; j < columnWidths.Length; j++ )
        {
          var cell = HelperFunctions.CreateTableCell();
          row.Add( cell );
        }

        newTable.Add( row );
      }
      return newTable;
    }

    /// <summary>
    /// Create and return a cell of a table        
    /// </summary>        
    internal static XElement CreateTableCell( double w = 2310 )
    {
      return new XElement( XName.Get( "tc", Document.w.NamespaceName ),
                           new XElement( XName.Get( "tcPr", Document.w.NamespaceName ),
                                         new XElement( XName.Get( "tcW", Document.w.NamespaceName ),
                                                       new XAttribute( XName.Get( "w", Document.w.NamespaceName ), w ),
                                                       new XAttribute( XName.Get( "type", Document.w.NamespaceName ), "dxa" ) ) ),
                           new XElement( XName.Get( "p", Document.w.NamespaceName ), new XElement( XName.Get( "pPr", Document.w.NamespaceName ) ) ) );
    }













    internal static void RenumberIDs( Document document )
    {
      IEnumerable<XAttribute> trackerIDs =
                      ( from d in document._mainDoc.Descendants()
                        where d.Name.LocalName == "ins" || d.Name.LocalName == "del"
                        select d.Attribute( XName.Get( "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main" ) ) );

      for( int i = 0; i < trackerIDs.Count(); i++ )
        trackerIDs.ElementAt( i ).Value = i.ToString();
    }

    internal static Paragraph GetFirstParagraphEffectedByInsert( Document document, int index )
    {
      // This document contains no Paragraphs and insertion is at index 0
      var docParagraphs = document.Paragraphs;
      if( docParagraphs.Count() == 0 && index == 0 )
        return null;

      foreach( Paragraph p in docParagraphs )
      {
        if( p._endIndex >= index )
          return p;
      }

      throw new ArgumentOutOfRangeException();
    }

    internal static List<XElement> FormatInput( string text, XElement rPr )
    {
      var newRuns = new List<XElement>();
      var tabRun = new XElement( Document.w + "tab" );
      var breakRun = new XElement( Document.w + "br" );

      var sb = new StringBuilder();

      if( string.IsNullOrEmpty( text ) )
      {
        return newRuns; //I dont wanna get an exception if text == null, so just return empy list
      }

      char lastCharacter = '\0';
      foreach( char c in text )
      {
        switch( c )
        {
          case '\t':
            if( sb.Length > 0 )
            {
              var t = new XElement( Document.w + "t", sb.ToString() );
              Xceed.Document.NET.Text.PreserveSpace( t );
              newRuns.Add( new XElement( Document.w + "r", rPr, t ) );
              sb = new StringBuilder();
            }
            newRuns.Add( new XElement( Document.w + "r", rPr, tabRun ) );
            break;
          case '\n':
            if( lastCharacter == '\r' )
              break;
            if( sb.Length > 0 )
            {
              var t = new XElement( Document.w + "t", sb.ToString() );
              Xceed.Document.NET.Text.PreserveSpace( t );
              newRuns.Add( new XElement( Document.w + "r", rPr, t ) );
              sb = new StringBuilder();
            }
            newRuns.Add( new XElement( Document.w + "r", rPr, breakRun ) );
            break;
          case '\r':
            if( sb.Length > 0 )
            {
              var t = new XElement( Document.w + "t", sb.ToString() );
              Xceed.Document.NET.Text.PreserveSpace( t );
              newRuns.Add( new XElement( Document.w + "r", rPr, t ) );
              sb = new StringBuilder();
            }
            newRuns.Add( new XElement( Document.w + "r", rPr, breakRun ) );
            break;

          default:
            if( !RestrictedXmlCharacters.Contains( c ) )
              sb.Append( c );
            break;
        }

        lastCharacter = c;
      }

      if( sb.Length > 0 )
      {
        var t = new XElement( Document.w + "t", sb.ToString() );
        Xceed.Document.NET.Text.PreserveSpace( t );
        newRuns.Add( new XElement( Document.w + "r", rPr, t ) );
      }

      return newRuns;
    }

    internal static XElement[] SplitParagraph( Paragraph p, int index )
    {
      // In this case edit dosent really matter, you have a choice.
      Run r = p.GetFirstRunEffectedByEdit( index, EditType.ins );

      XElement[] split;
      XElement before, after;

      if( r.Xml.Parent.Name.LocalName == "ins" )
      {
        split = p.SplitEdit( r.Xml.Parent, index, EditType.ins );
        before = new XElement( p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[ 0 ] );
        after = new XElement( p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[ 1 ] );
      }

      else if( r.Xml.Parent.Name.LocalName == "del" )
      {
        split = p.SplitEdit( r.Xml.Parent, index, EditType.del );

        before = new XElement( p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[ 0 ] );
        after = new XElement( p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[ 1 ] );
      }

      else
      {
        split = Run.SplitRun( r, index );

        before = new XElement( p.Xml.Name, p.Xml.Attributes(), r.Xml.ElementsBeforeSelf(), split[ 0 ] );
        after = new XElement( p.Xml.Name, p.Xml.Attributes(), split[ 1 ], r.Xml.ElementsAfterSelf() );
      }

      if( before.Elements().Count() == 0 )
        before = null;

      if( after.Elements().Count() == 0 )
        after = null;

      return new XElement[] { before, after };
    }

    /// <!-- 
    /// Bug found and fixed by trnilse. To see the change, 
    /// please compare this release to the previous release using TFS compare.
    /// -->
    internal static bool IsSameFile( Stream streamOne, Stream streamTwo )
    {
      int file1byte, file2byte;

      if( streamOne.Length != streamTwo.Length )
      {
        // Return false to indicate files are different
        return false;
      }

      // Read and compare a byte from each file until either a
      // non-matching set of bytes is found or until the end of
      // file1 is reached.
      do
      {
        // Read one byte from each file.
        file1byte = streamOne.ReadByte();
        file2byte = streamTwo.ReadByte();
      }
      while( ( file1byte == file2byte ) && ( file1byte != -1 ) );

      // Return the success of the comparison. "file1byte" is 
      // equal to "file2byte" at this point only if the files are 
      // the same.

      streamOne.Position = 0;
      streamTwo.Position = 0;

      return ( ( file1byte - file2byte ) == 0 );
    }

    /// <summary> 
    /// Add the default numbering.xml if it is missing from this document
    /// </summary> 
    internal static XDocument AddDefaultNumberingXml( Package package )
    {
      XDocument numberingDoc;

      var numberingPart = package.CreatePart( new Uri( "/word/numbering.xml", UriKind.Relative ), "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", CompressionOption.Maximum );
      numberingDoc = DecompressXMLResource( HelperFunctions.GetResources( ResourceType.Numbering ) );

      using( TextWriter tw = new StreamWriter( new PackagePartStream( numberingPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        numberingDoc.Save( tw, SaveOptions.None );
      }

      var mainDocPart = GetMainDocumentPart( package );

      mainDocPart.CreateRelationship( numberingPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" );
      return numberingDoc;
    }

    internal static List CreateItemInList( List list, 
                                           string listText, 
                                           int level = 0,
                                           ListItemType listType = ListItemType.Numbered,
                                           int? startNumber = null, 
                                           bool trackChanges = false, 
                                           bool continueNumbering = false, 
                                           Formatting formatting = null )
    {
      if( list.NumId == 0 )
      {
        list.CreateNewNumberingNumId( level, listType, startNumber, continueNumbering );
      }

      if( listText != null )
      {
        var newSection = new XElement( XName.Get( "p", Document.w.NamespaceName ) );
        var last_pPr = ( list.Items != null ) && ( list.Items.Count > 0 ) ? list.Items.Last().GetOrCreate_pPr() : null;
        // This is the first listItem.
        if( last_pPr == null )
        {
          newSection.Add( new XElement( XName.Get( "pPr", Document.w.NamespaceName ),
                          new XElement( XName.Get( "numPr", Document.w.NamespaceName ),
                          new XElement( XName.Get( "ilvl", Document.w.NamespaceName ), new XAttribute( Document.w + "val", level ) ),
                          new XElement( XName.Get( "numId", Document.w.NamespaceName ), new XAttribute( Document.w + "val", list.NumId ) ) ) ) );
        }
        // Use the Paragraph properties of the last ListItem.
        else
        {
          newSection.Add( new XElement( last_pPr ) );
          // Use the wanted level.
          var ilvl = newSection.Descendants( XName.Get( "ilvl", Document.w.NamespaceName ) ).FirstOrDefault();
          if( ilvl != null )
          {
            ilvl.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), level );
          }
        }

        if( formatting == null )
        {
          newSection.Add( new XElement( XName.Get( "r", Document.w.NamespaceName ), new XElement( XName.Get( "t", Document.w.NamespaceName ), listText ) ) );
        }
        else
        {
          newSection.Add( HelperFunctions.FormatInput( listText, formatting.Xml ) );
        }

        if( trackChanges )
          newSection = CreateEdit( EditType.ins, DateTime.Now, newSection );

        if( startNumber == null )
          list.AddItem( new Paragraph( list.Document, newSection, 0, ContainerType.Paragraph ) );
        else
          list.AddItemWithStartValue( new Paragraph( list.Document, newSection, 0, ContainerType.Paragraph ), (int)startNumber );
      }
      return list;
    }






    internal static UnderlineStyle GetUnderlineStyle( string underlineStyle )
    {
      switch( underlineStyle )
      {
        case "single":
          return UnderlineStyle.singleLine;
        case "double":
          return UnderlineStyle.doubleLine;
        case "thick":
          return UnderlineStyle.thick;
        case "dotted":
          return UnderlineStyle.dotted;
        case "dottedHeavy":
          return UnderlineStyle.dottedHeavy;
        case "dash":
          return UnderlineStyle.dash;
        case "dashedHeavy":
          return UnderlineStyle.dashedHeavy;
        case "dashLong":
          return UnderlineStyle.dashLong;
        case "dashLongHeavy":
          return UnderlineStyle.dashLongHeavy;
        case "dotDash":
          return UnderlineStyle.dotDash;
        case "dashDotHeavy":
          return UnderlineStyle.dashDotHeavy;
        case "dotDotDash":
          return UnderlineStyle.dotDotDash;
        case "dashDotDotHeavy":
          return UnderlineStyle.dashDotDotHeavy;
        case "wave":
          return UnderlineStyle.wave;
        case "wavyHeavy":
          return UnderlineStyle.wavyHeavy;
        case "wavyDouble":
          return UnderlineStyle.wavyDouble;
        case "words":
          return UnderlineStyle.words;
        default:
          return UnderlineStyle.none;
      }
    }

    internal static bool ContainsEveryChildOf( XElement elementWanted, XElement elementToValidate, MatchFormattingOptions formattingOptions )
    {
      foreach( XElement subElement in elementWanted.Elements() )
      {
        if( !elementToValidate.Elements( subElement.Name ).Where( bElement => bElement.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) == subElement.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) ).Any() )
          return false;
      }

      if( formattingOptions == MatchFormattingOptions.ExactMatch )
        return elementWanted.Elements().Count() == elementToValidate.Elements().Count();

      return true;
    }

    internal static Color GetColorFromHtml( string stringColor, string autoColor = "FFFFFF" )
    {
      Debug.Assert( !string.IsNullOrEmpty( stringColor ), "stringColor should not be null or empty." );
      // Default to White when "auto".
      if( stringColor == "auto" )
      {
        stringColor = autoColor;
      }
      Debug.Assert( stringColor.Length == 6, "stringColor should have a length of 6 characters." );

      return Color.FromArgb( Convert.ToInt32( stringColor.Substring( 0, 2 ), 16 ), Convert.ToInt32( stringColor.Substring( 2, 2 ), 16 ), Convert.ToInt32( stringColor.Substring( 4, 2 ), 16 ) );
    }

    internal static string GetListItemType( Paragraph p, Document document )
    {
      var paragraphNumberPropertiesDescendants = p.ParagraphNumberProperties.Descendants();
      var ilvlNode = paragraphNumberPropertiesDescendants.FirstOrDefault( el => el.Name.LocalName == "ilvl" );
      var ilvlValue = ( ilvlNode != null ) ? ilvlNode.Attribute( Document.w + "val" ).Value : null;

      var numIdNode = paragraphNumberPropertiesDescendants.FirstOrDefault( el => el.Name.LocalName == "numId" );
      var numIdValue = ( numIdNode != null ) ? numIdNode.Attribute( Document.w + "val" ).Value : null;

      var abstractNumNode = HelperFunctions.GetAbstractNum( document, numIdValue );
      if( abstractNumNode != null )
      {
        // Find lvl node.
        var lvlNodes = abstractNumNode.Descendants().Where( n => n.Name.LocalName == "lvl" );
        // No lvl, check if a numStyleLink is used.
        if( lvlNodes.Count() == 0 )
        {
          var linkedStyleNumId = HelperFunctions.GetLinkedStyleNumId( document, numIdValue );
          if( linkedStyleNumId != -1 )
          {
            abstractNumNode = HelperFunctions.GetAbstractNum( document, linkedStyleNumId.ToString() );
            if( abstractNumNode != null )
            {
              lvlNodes = abstractNumNode.Descendants().Where( n => n.Name.LocalName == "lvl" );
            }
          }
        }
        XElement lvlNode = null;
        foreach( XElement node in lvlNodes )
        {
          if( node.Attribute( Document.w + "ilvl" ).Value.Equals( ilvlValue ) )
          {
            lvlNode = node;
            break;
          }
          else if( ilvlValue == null )
          {
            var numStyleNode = node.Descendants().FirstOrDefault( n => n.Name.LocalName == "pStyle" );
            if( ( numStyleNode != null ) && numStyleNode.GetAttribute( Document.w + "val" ).Equals( p.StyleId ) )
            {
              lvlNode = node;
              break;
            }
          }
        }

        if( lvlNode != null )
        {
          var numFmtNode = lvlNode.Descendants().FirstOrDefault( n => n.Name.LocalName == "numFmt" );
          if( numFmtNode != null )
            return numFmtNode.Attribute( Document.w + "val" ).Value;
        }
      }

      return null;
    }

    internal static XElement GetAbstractNum( Document document, string numId )
    {
      if( document == null )
        return null;
      if( numId == null )
        return null;

      var abstractNumNodeValue = HelperFunctions.GetAbstractNumIdValue( document, numId );
      if( string.IsNullOrEmpty( abstractNumNodeValue ) )
        return null;

      // Find abstractNum node in numbering.
      var documentNumberingDescendants = document._numbering.Descendants();
      var abstractNumNodes = documentNumberingDescendants.Where( n => n.Name.LocalName == "abstractNum" );
      var abstractNumNode = abstractNumNodes.FirstOrDefault( node => node.Attribute( Document.w + "abstractNumId" ).Value.Equals( abstractNumNodeValue ) );

      return abstractNumNode;
    }

    internal static string GetAbstractNumIdValue( Document document, string numId )
    {
      if( document == null )
        return null;
      if( numId == null )
        return null;

      // Find num node in numbering.
      var documentNumberingDescendants = document._numbering.Descendants();
      var numNodes = documentNumberingDescendants.Where( n => n.Name.LocalName == "num" );
      var numNode = numNodes.FirstOrDefault( node => node.Attribute( Document.w + "numId" ).Value.Equals( numId ) );
      if( numNode == null )
        return null;

      // Get abstractNumId node and its value from numNode.
      var abstractNumIdNode = numNode.Descendants().FirstOrDefault( n => n.Name.LocalName == "abstractNumId" );
      if( abstractNumIdNode == null )
        return null;
      var abstractNumNodeValue = abstractNumIdNode.Attribute( Document.w + "val" ).Value;
      if( string.IsNullOrEmpty( abstractNumNodeValue ) )
        return null;

      return abstractNumNodeValue;
    }

    internal static string GetListItemStartValue( List list, int level )
    {
      var abstractNumElement = list.GetAbstractNum( list.NumId );
      if( abstractNumElement == null )
        return "1";

      //Find lvl node
      var lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
      var lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( Document.w + "ilvl" ).Equals( level.ToString() ) );
      // No ilvl, check if a numStyleLink is used.
      if( lvlNode == null )
      {
        var linkedStyleNumId = HelperFunctions.GetLinkedStyleNumId( list.Document, list.NumId.ToString() );
        if( linkedStyleNumId != -1 )
        {
          abstractNumElement = list.GetAbstractNum( linkedStyleNumId );
          if( abstractNumElement == null )
            return "1";
          lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
          lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( Document.w + "ilvl" ).Equals( level.ToString() ) );
        }
        if( lvlNode == null )
          return "1";
      }

      var startNode = lvlNode.Descendants().FirstOrDefault( n => n.Name.LocalName == "start" );
      if( startNode == null )
        return "1";
      return startNode.GetAttribute( Document.w + "val" );
    }

    internal static string GetListItemTextFormat( List list, int level, out Formatting formatting )
    {
      formatting = null;
      var abstractNumElement = list.GetAbstractNum( list.NumId );
      if( abstractNumElement == null )
        return "%1.";

      //Find lvl node
      var lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
      var lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( Document.w + "ilvl" ).Equals( level.ToString() ) );
      // No ilvl, check if a numStyleLink is used.
      if( lvlNode == null )
      {
        var linkedStyleNumId = HelperFunctions.GetLinkedStyleNumId( list.Document, list.NumId.ToString() );
        if( linkedStyleNumId != -1 )
        {
          abstractNumElement = list.GetAbstractNum( linkedStyleNumId );
          if( abstractNumElement == null )
            return "%1.";
          lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
          lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( Document.w + "ilvl" ).Equals( level.ToString() ) );
        }
        if( lvlNode == null )
          return "%1.";
      }

      formatting = Formatting.Parse( lvlNode.Descendants().FirstOrDefault( n => n.Name.LocalName == "rPr" ) );

      var textFormatNode = lvlNode.Descendants().FirstOrDefault( n => n.Name.LocalName == "lvlText" );
      if( textFormatNode == null )
        return "%1.";
      return textFormatNode.GetAttribute( Document.w + "val" );
    }

    internal static XElement GetListItemAlignment( List list, int level )
    {
      var abstractNumElement = list.GetAbstractNum( list.NumId );
      if( abstractNumElement == null )
        return null;

      //Find lvl node
      var lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
      var lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( Document.w + "ilvl" ).Equals( level.ToString() ) );
      // No ilvl, check if a numStyleLink is used.
      if( lvlNode == null )
      {
        var linkedStyleNumId = HelperFunctions.GetLinkedStyleNumId( list.Document, list.NumId.ToString() );
        if( linkedStyleNumId != -1 )
        {
          abstractNumElement = list.GetAbstractNum( linkedStyleNumId );
          if( abstractNumElement == null )
            return null;
          lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
          lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( Document.w + "ilvl" ).Equals( level.ToString() ) );
        }
        if( lvlNode == null )
          return null;
      }

      var pPr = lvlNode.Descendants().FirstOrDefault( n => n.Name.LocalName == "pPr" );
      if( pPr != null )
      {
        var ind = pPr.Descendants().FirstOrDefault( n => n.Name.LocalName == "ind" );
        if( ind != null )
        {
          return ind;
        }
      }
      return null;
    }

    internal static Border GetBorderFromXml( XElement xml )
    {
      if( xml == null )
        return null;

      var borderSize = BorderSize.one;
      var borderColor = Color.Black;
      var borderSpace = 0f;
      var borderStyle = BorderStyle.Tcbs_single;

      var bdrColor = xml.Attribute( XName.Get( "color", Document.w.NamespaceName ) );
      if( ( bdrColor != null ) && ( bdrColor.Value != "auto" ) )
      {
        borderColor = HelperFunctions.GetColorFromHtml( bdrColor.Value );
      }
      var size = xml.Attribute( XName.Get( "sz", Document.w.NamespaceName ) );
      if( size != null )
      {
        var sizeValue = System.Convert.ToSingle( size.Value );
        if( sizeValue == 2 )
          borderSize = BorderSize.one;
        else if( sizeValue == 4 )
          borderSize = BorderSize.two;
        else if( sizeValue == 6 )
          borderSize = BorderSize.three;
        else if( sizeValue == 8 )
          borderSize = BorderSize.four;
        else if( sizeValue == 12 )
          borderSize = BorderSize.five;
        else if( sizeValue == 18 )
          borderSize = BorderSize.six;
        else if( sizeValue == 24 )
          borderSize = BorderSize.seven;
        else if( sizeValue == 36 )
          borderSize = BorderSize.eight;
        else if( sizeValue == 48 )
          borderSize = BorderSize.nine;
        else
          borderSize = BorderSize.one;
      }
      var space = xml.Attribute( XName.Get( "space", Document.w.NamespaceName ) );
      if( space != null )
      {
        borderSpace = System.Convert.ToSingle( space.Value );
      }
      var bdrStyle = xml.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
      if( bdrStyle != null )
      {
        borderStyle = (BorderStyle)Enum.Parse( typeof( BorderStyle ), "Tcbs_" + bdrStyle.Value );
      }

      return new Border( borderStyle, borderSize, borderSpace, borderColor );
    }

    internal static void UpdateParagraphFromStyledParagraph( Paragraph p, Paragraph styledParagraph )
    {
      if( ( p == null ) || ( styledParagraph == null ) )
        return;

      var paragraph_Ppr = p.GetOrCreate_pPr();
      var styleParagraph_pPr = styledParagraph.GetOrCreate_pPr();
      // Loop through each styled paragraph properties to add the ones not currently in paragraph properties.
      foreach( var styleParagraphElement in styleParagraph_pPr.Elements() )
      {
        var paragraphElement = paragraph_Ppr.Element( styleParagraphElement.Name );
        // paragraph doesn't contains the styled property, add it.
        if( paragraphElement == null )
        {
          paragraph_Ppr.Add( styleParagraphElement );
        }
        else
        {
          // Add missing "tab" from styled paragraph "tabs".
          if( paragraphElement.Name.LocalName == "tabs" )
          {
            var styledTabs = styleParagraphElement.Elements();
            foreach( var styledTab in styledTabs )
            {
              if( paragraphElement.Elements( XName.Get( "tab", Document.w.NamespaceName ) )
                                  .FirstOrDefault( x => x.GetAttribute( XName.Get( "pos", Document.w.NamespaceName ) ) == styledTab.GetAttribute( XName.Get( "pos", Document.w.NamespaceName ) ) ) == null )
              {
                paragraphElement.Add( styledTab );
              }
            }
          }
          // Paragraph contains the property, add the missing attributes from this property.
          foreach( var att in styleParagraphElement.Attributes() )
          {
            if( paragraphElement.Attribute( att.Name ) == null && p.CanAddAttribute( att ) )
            {
              paragraphElement.Add( att );
            }
          }
        }
      }

      var paragraph_rPr = p.GetOrCreate_rPr();
      var styleParagraph_rPr = styledParagraph.GetOrCreate_rPr();
      // Loop through each styled paragraph run properties to add the ones not currently in paragraph run properties.
      foreach( var styleParagraphElement in styleParagraph_rPr.Elements() )
      {
        var runElement = paragraph_rPr.Element( styleParagraphElement.Name );
        // paragraph doesn't contains the styled property, add it.
        if( runElement == null )
        {
          paragraph_rPr.Add( styleParagraphElement );
        }
        else
        {
          // Paragraph contains the property, add the missing attributes from this property.
          foreach( var att in styleParagraphElement.Attributes() )
          {
            if( runElement.Attribute( att.Name ) == null )
            {
              runElement.Add( att );
            }
          }
        }
      }
      // Reset Backers because the paragraph XML content may have changed with the syled elements added.
      p.ResetBackers();
    }

    internal static XElement GetParagraphStyleFromStyleId( Document document, string styleIdToFind )
    {
      if( ( document == null ) || string.IsNullOrEmpty( styleIdToFind ) )
        return null;

      var paragraphStyles =
        (
            from s in document._styles.Element( Document.w + "styles" ).Elements( Document.w + "style" )
            let type = s.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
            where ( ( type != null ) && ( type.Value == "paragraph" ) )
            select s
        );

      // Check if this Paragraph styleIdToFind exists in _styles.
      var currentParagraphStyle =
     (
         from s in paragraphStyles
         let styleId = s.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) )
         where ( ( styleId != null ) && ( styleId.Value == styleIdToFind ) )
         select s
     ).FirstOrDefault();

      return currentParagraphStyle;
    }

    internal static XElement GetParagraphStyleFromStyleName( Document document, string styleNameToFind )
    {
      if( ( document == null ) || string.IsNullOrEmpty( styleNameToFind ) )
        return null;

      var paragraphStyles =
        (
            from s in document._styles.Element( Document.w + "styles" ).Elements( Document.w + "style" )
            let type = s.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
            where ( ( type != null ) && ( type.Value == "paragraph" ) )
            select s
        );

      // Check if this Paragraph styleNameToFind exists in _styles.
      var currentParagraphStyle = paragraphStyles.FirstOrDefault( x => ( x.Element( XName.Get( "name", Document.w.NamespaceName ) ) != null )
                                                                    && ( x.Element( XName.Get( "name", Document.w.NamespaceName ) ).Attribute( XName.Get( "val", Document.w.NamespaceName ) ) != null )
                                                                    && ( x.Element( XName.Get( "name", Document.w.NamespaceName ) ).Attribute( XName.Get( "val", Document.w.NamespaceName ) ).Value.ToLower().Equals( styleNameToFind.ToLower() ) ) );


      return currentParagraphStyle;
    }

    internal static void CopyStream( Stream input, Stream output, int bufferSize = 32768 )
    {
      byte[] buffer = new byte[ bufferSize ];
      int read;
      while( ( read = input.Read( buffer, 0, buffer.Length ) ) > 0 )
      {
        output.Write( buffer, 0, read );
      }
    }

    internal static XElement GetStyle( Document fileToConvert, string styleId )
    {
      if( fileToConvert == null )
        throw new ArgumentNullException( "fileToConvert" );
      if( string.IsNullOrEmpty( styleId ) )
        throw new ArgumentNullException( "styleId" );

      var styles = fileToConvert._styles.Element( XName.Get( "styles", Document.w.NamespaceName ) );
      return styles.Elements( XName.Get( "style", Document.w.NamespaceName ) )
                   .FirstOrDefault( x => ( x.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) ) != null ) && ( x.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) ).Value == styleId ) );
    }

    internal static bool TryParseFloat( string s, out float result )
    {
      if( float.TryParse( s, NumberStyles.Any, CultureInfo.InvariantCulture, out result ) )
        return true;
      return false;
    }

    internal static bool TryParseDouble( string s, out double result )
    {
      if( double.TryParse( s, NumberStyles.Any, CultureInfo.InvariantCulture, out result ) )
        return true;
      return false;
    }

    internal static bool TryParseInt( string s, out int result )
    {
      if( int.TryParse( s, NumberStyles.Any, CultureInfo.InvariantCulture, out result ) )
        return true;
      return false;
    }

    private static int GetLinkedStyleNumId( Document document, string numId )
    {
      Debug.Assert( document != null, "document should not be null" );

      var abstractNumElement = HelperFunctions.GetAbstractNum( document, numId );
      if( abstractNumElement != null )
      {
        var numStyleLink = abstractNumElement.Element( XName.Get( "numStyleLink", Document.w.NamespaceName ) );
        if( numStyleLink != null )
        {
          var val = numStyleLink.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
          if( !string.IsNullOrEmpty( val.Value ) )
          {
            var linkedStyle = HelperFunctions.GetStyle( document, val.Value );
            if( linkedStyle != null )
            {
              var linkedNumId = linkedStyle.Descendants( XName.Get( "numId", Document.w.NamespaceName ) ).FirstOrDefault();
              if( linkedNumId != null )
              {
                var linkedNumIdVal = linkedNumId.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
                if( !string.IsNullOrEmpty( linkedNumIdVal.Value ) )
                  return Int32.Parse( linkedNumIdVal.Value );
              }
            }
          }
        }
      }

      return -1;
    }
  }
}
