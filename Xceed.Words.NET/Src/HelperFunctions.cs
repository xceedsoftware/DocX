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
using System.IO.Packaging;
using System.Xml.Linq;
using System.IO;
using System.Reflection;
using System.IO.Compression;
using System.Security.Principal;
using System.Globalization;
using System.Xml;

namespace Xceed.Words.NET
{
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

    internal static void CreateRelsPackagePart( DocX Document, Uri uri )
    {
      PackagePart pp = Document._package.CreatePart( uri, DocX.ContentTypeApplicationRelationShipXml, CompressionOption.Maximum );
      using( TextWriter tw = new StreamWriter( new PackagePartStream( pp.GetStream() ) ) )
      {
        XDocument d = new XDocument
        (
            new XDeclaration( "1.0", "UTF-8", "yes" ),
            new XElement( XName.Get( "Relationships", DocX.rel.NamespaceName ) )
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
          return (Xml.Parent.Name.LocalName != "tabs" ) ? 1 : 0;
        case "br":
          return 1;
        case "t":
          goto case "delText";
        case "delText":
          return Xml.Value.Length;
        case "tr":
          goto case "br";
        case "tc":
          goto case "br";
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

      if( Xml.HasElements )
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

      // e is a w:t element, it must exist inside a w:r element or a w:tabs, lets climb until we find it.
      while( (e != null) && !e.Name.Equals( XName.Get( "r", DocX.w.NamespaceName ) ) && !e.Name.Equals( XName.Get( "tabs", DocX.w.NamespaceName ) ) )
        e = e.Parent;

      FormattedText ft = new FormattedText();
      ft.text = text;
      ft.index = 0;
      ft.formatting = null;

      if( e != null )
      {
        // e is a w:r element, lets find the rPr element.
        XElement rPr = e.Element( XName.Get( "rPr", DocX.w.NamespaceName ) );

        // Return text with formatting.
        if( rPr != null )
          ft.formatting = Formatting.Parse( rPr );
      }

      return ft;
    }

    internal static string ToText( XElement e )
    {
      switch( e.Name.LocalName )
      {
        case "tab":
          // Do not add "\t" for TabStopPositions defined in "tabs".
          return ((e.Parent != null) && e.Parent.Name.Equals( XName.Get( "tabs", DocX.w.NamespaceName ) )) ? "" : "\t";
        case "br":
          return "\n";
        case "t":
          goto case "delText";
        case "delText":
          {


            return e.Value;
          }
        case "tr":
          goto case "br";
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

        var themeFontLang = settings.Root.Element( XName.Get( "themeFontLang", DocX.w.NamespaceName ) );
        themeFontLang.SetAttributeValue( XName.Get( "val", DocX.w.NamespaceName ), CultureInfo.CurrentCulture );

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

    internal static void CreateCustomPropertiesPart( DocX document )
    {
      var customPropertiesPart = document._package.CreatePart( new Uri( "/docProps/custom.xml", UriKind.Relative ), "application/vnd.openxmlformats-officedocument.custom-properties+xml", CompressionOption.Maximum );

      var customPropDoc = new XDocument
      (
          new XDeclaration( "1.0", "UTF-8", "yes" ),
          new XElement
          (
              XName.Get( "Properties", DocX.customPropertiesSchema.NamespaceName ),
              new XAttribute( XNamespace.Xmlns + "vt", DocX.customVTypesSchema )
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

      stylesDoc = HelperFunctions.DecompressXMLResource( "Xceed.Words.NET.Resources.default_styles.xml.gz" );
      var lang = stylesDoc.Root.Element( XName.Get( "docDefaults", DocX.w.NamespaceName ) ).Element( XName.Get( "rPrDefault", DocX.w.NamespaceName ) ).Element( XName.Get( "rPr", DocX.w.NamespaceName ) ).Element( XName.Get( "lang", DocX.w.NamespaceName ) );
      lang.SetAttributeValue( XName.Get( "val", DocX.w.NamespaceName ), CultureInfo.CurrentCulture );

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

      return
      (
          new XElement( DocX.w + t.ToString(),
              new XAttribute( DocX.w + "id", 0 ),
              new XAttribute( DocX.w + "author", WindowsIdentity.GetCurrent().Name ),
              new XAttribute( DocX.w + "date", edit_time ),
          content )
      );
    }

    internal static XElement CreateTable( int rowCount, int columnCount )
    {
      if( ( rowCount <= 0 ) || ( columnCount <= 0 ) )
      {
        throw new ArgumentOutOfRangeException( "Row and Column count must be greater than 0." );
      }

      int[] columnWidths = new int[columnCount];
      for (int i = 0; i < columnCount; i++)
      {
        columnWidths[i] = 2310;
      }
      return CreateTable(rowCount, columnWidths);
    }

    internal static XElement CreateTable( int rowCount, int[] columnWidths )
    {
      var newTable = new XElement( XName.Get( "tbl", DocX.w.NamespaceName ),
                                   new XElement( XName.Get( "tblPr", DocX.w.NamespaceName ),
                                                 new XElement( XName.Get( "tblStyle", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), "TableGrid" ) ),
                                                 new XElement( XName.Get( "tblW", DocX.w.NamespaceName ), new XAttribute( XName.Get( "w", DocX.w.NamespaceName ), "5000" ), new XAttribute( XName.Get( "type", DocX.w.NamespaceName ), "auto" ) ),
                                                 new XElement( XName.Get( "tblLook", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), "04A0" ) ) ) );

      for( int i = 0; i < rowCount; i++ )
      {
        var row = new XElement( XName.Get( "tr", DocX.w.NamespaceName ) );

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
      return new XElement( XName.Get( "tc", DocX.w.NamespaceName ),
                           new XElement( XName.Get( "tcPr", DocX.w.NamespaceName ),
                                         new XElement( XName.Get( "tcW", DocX.w.NamespaceName ),
                                                       new XAttribute( XName.Get( "w", DocX.w.NamespaceName ), w ),
                                                       new XAttribute( XName.Get( "type", DocX.w.NamespaceName ), "dxa" ) ) ),
                           new XElement( XName.Get( "p", DocX.w.NamespaceName ), new XElement( XName.Get( "pPr", DocX.w.NamespaceName ) ) ) );
    }

    internal static void RenumberIDs( DocX document )
    {
      IEnumerable<XAttribute> trackerIDs =
                      ( from d in document._mainDoc.Descendants()
                        where d.Name.LocalName == "ins" || d.Name.LocalName == "del"
                        select d.Attribute( XName.Get( "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main" ) ) );

      for( int i = 0; i < trackerIDs.Count(); i++ )
        trackerIDs.ElementAt( i ).Value = i.ToString();
    }

    internal static Paragraph GetFirstParagraphEffectedByInsert( DocX document, int index )
    {
      // This document contains no Paragraphs and insertion is at index 0
      if( document.Paragraphs.Count() == 0 && index == 0 )
        return null;

      foreach( Paragraph p in document.Paragraphs )
      {
        if( p._endIndex >= index )
          return p;
      }

      throw new ArgumentOutOfRangeException();
    }

    internal static List<XElement> FormatInput( string text, XElement rPr )
    {
      var newRuns = new List<XElement>();
      var tabRun = new XElement( DocX.w + "tab" );
      var breakRun = new XElement( DocX.w + "br" );

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
              var t = new XElement( DocX.w + "t", sb.ToString() );
              Xceed.Words.NET.Text.PreserveSpace( t );
              newRuns.Add( new XElement( DocX.w + "r", rPr, t ) );
              sb = new StringBuilder();
            }
            newRuns.Add( new XElement( DocX.w + "r", rPr, tabRun ) );
            break;
          case '\n':
            if( lastCharacter == '\r' )
              break;
            if( sb.Length > 0 )
            {
              var t = new XElement( DocX.w + "t", sb.ToString() );
              Xceed.Words.NET.Text.PreserveSpace( t );
              newRuns.Add( new XElement( DocX.w + "r", rPr, t ) );
              sb = new StringBuilder();
            }
            newRuns.Add( new XElement( DocX.w + "r", rPr, breakRun ) );
            break;
          case '\r':
            if( sb.Length > 0 )
            {
              var t = new XElement( DocX.w + "t", sb.ToString() );
              Xceed.Words.NET.Text.PreserveSpace( t );
              newRuns.Add( new XElement( DocX.w + "r", rPr, t ) );
              sb = new StringBuilder();
            }
            newRuns.Add( new XElement( DocX.w + "r", rPr, breakRun ) );
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
        var t = new XElement( DocX.w + "t", sb.ToString() );
        Xceed.Words.NET.Text.PreserveSpace( t );
        newRuns.Add( new XElement( DocX.w + "r", rPr, t ) );
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
    internal static XDocument AddDefaultNumberingXml(Package package)
    {
      XDocument numberingDoc;

      var numberingPart = package.CreatePart(new Uri("/word/numbering.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", CompressionOption.Maximum);
      numberingDoc = DecompressXMLResource("Xceed.Words.NET.Resources.numbering.xml.gz");

      using( TextWriter tw = new StreamWriter( new PackagePartStream( numberingPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        numberingDoc.Save( tw, SaveOptions.None );
      }

      var mainDocPart = GetMainDocumentPart( package );

      mainDocPart.CreateRelationship(numberingPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
      return numberingDoc;
    }

    internal static List CreateItemInList(List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false)
    {
        if (list.NumId == 0)
        {
            list.CreateNewNumberingNumId(level, listType, startNumber, continueNumbering);
        }

        if (listText != null)
        {
            var newSection = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName),
                new XElement(XName.Get("pPr", DocX.w.NamespaceName),
                new XElement(XName.Get("numPr", DocX.w.NamespaceName),
                new XElement(XName.Get("ilvl", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", level)),
                new XElement(XName.Get("numId", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", list.NumId)))),
                new XElement(XName.Get("r", DocX.w.NamespaceName), new XElement(XName.Get("t", DocX.w.NamespaceName), listText))
            );

            if (trackChanges)
                newSection = CreateEdit(EditType.ins, DateTime.Now, newSection);

            if (startNumber == null)
                list.AddItem(new Paragraph(list.Document, newSection, 0, ContainerType.Paragraph));
            else
                list.AddItemWithStartValue(new Paragraph(list.Document, newSection, 0, ContainerType.Paragraph), (int)startNumber);
        }
        return list;
    }

    internal static UnderlineStyle GetUnderlineStyle(string styleName)
    {
        switch (styleName)
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

    internal static bool ContainsEveryChildOf(XElement elementWanted, XElement elementToValidate, MatchFormattingOptions formattingOptions)
    {
        foreach (XElement subElement in elementWanted.Elements())
        {
            if (!elementToValidate.Elements(subElement.Name).Where(bElement => bElement.GetAttribute(XName.Get("val", DocX.w.NamespaceName)) == subElement.GetAttribute(XName.Get("val", DocX.w.NamespaceName))).Any())
                return false;
        }

        if (formattingOptions == MatchFormattingOptions.ExactMatch)
            return elementWanted.Elements().Count() == elementToValidate.Elements().Count();

        return true;
    }

    internal static string GetListItemType( Paragraph p, DocX document )
    {
      var paragraphNumberPropertiesDescendants = p.ParagraphNumberProperties.Descendants();
      var ilvlNode = paragraphNumberPropertiesDescendants.FirstOrDefault( el => el.Name.LocalName == "ilvl" );
      var ilvlValue = ( ilvlNode != null) ? ilvlNode.Attribute( DocX.w + "val" ).Value : null;

      var numIdNode = paragraphNumberPropertiesDescendants.FirstOrDefault( el => el.Name.LocalName == "numId" );
      var numIdValue = ( numIdNode != null ) ? numIdNode.Attribute( DocX.w + "val" ).Value : null;

      //find num node in numbering 
      var documentNumberingDescendants = document._numbering.Descendants();
      var numNodes = documentNumberingDescendants.Where( n => n.Name.LocalName == "num" );
      XElement numNode = numNodes.FirstOrDefault( node => node.Attribute( DocX.w + "numId" ).Value.Equals( numIdValue ) );

      if( numNode != null )
      {
        //Get abstractNumId node and its value from numNode
        var abstractNumIdNode = numNode.Descendants().First( n => n.Name.LocalName == "abstractNumId" );
        var abstractNumNodeValue = abstractNumIdNode.Attribute( DocX.w + "val" ).Value;

        var abstractNumNodes = documentNumberingDescendants.Where( n => n.Name.LocalName == "abstractNum" );
        XElement abstractNumNode = abstractNumNodes.FirstOrDefault( node => node.Attribute( DocX.w + "abstractNumId" ).Value.Equals( abstractNumNodeValue ) );

        //Find lvl node
        var lvlNodes = abstractNumNode.Descendants().Where( n => n.Name.LocalName == "lvl" );
        XElement lvlNode = null;
        foreach( XElement node in lvlNodes )
        {
          if( node.Attribute( DocX.w + "ilvl" ).Value.Equals( ilvlValue ) )
          {
            lvlNode = node;
            break;
          }
          else if( ilvlValue == null )
          {
            var numStyleNode = node.Descendants().FirstOrDefault( n => n.Name.LocalName == "pStyle" );
            if( ( numStyleNode != null) && numStyleNode.GetAttribute( DocX.w + "val" ).Equals( p.StyleName ) )
            {
              lvlNode = node;
              break;
            }
          }
        }

        if( lvlNode != null )
        {
          var numFmtNode = lvlNode.Descendants().First( n => n.Name.LocalName == "numFmt" );
          return numFmtNode.Attribute( DocX.w + "val" ).Value;
        }
      }

      return null;
    }

    internal static string GetListItemStartValue( List list, int level )
    {
      var abstractNumElement = list.GetAbstractNum( list.NumId );

      //Find lvl node
      var lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
      var lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( DocX.w + "ilvl" ).Equals( level.ToString() ) );

      var startNode = lvlNode.Descendants().First( n => n.Name.LocalName == "start" );
      return startNode.GetAttribute( DocX.w + "val" );
    }

    internal static string GetListItemTextFormat( List list, int level )
    {
      var abstractNumElement = list.GetAbstractNum( list.NumId );

      //Find lvl node
      var lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
      var lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( DocX.w + "ilvl" ).Equals( level.ToString() ) );

      var textFormatNode = lvlNode.Descendants().First( n => n.Name.LocalName == "lvlText" );
      return textFormatNode.GetAttribute( DocX.w + "val" );
    }

    internal static XElement GetListItemAlignment( List list, int level )
    {
      var abstractNumElement = list.GetAbstractNum( list.NumId );

      //Find lvl node
      var lvlNodes = abstractNumElement.Descendants().Where( n => n.Name.LocalName == "lvl" );
      var lvlNode = lvlNodes.FirstOrDefault( n => n.GetAttribute( DocX.w + "ilvl" ).Equals( level.ToString() ) );

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
  }
}
