/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Represents a table of contents in the document
  /// </summary>
  public class TableOfContents : DocXElement
  {
    #region Private Constants

    private const string HeaderStyle = "TOCHeading";
    private const int RightTabPos = 9350;

    #endregion

    #region Internal Methods

    internal static TableOfContents CreateTableOfContents( DocX document, string title, TableOfContentsSwitches switches, string headerStyle = null, int lastIncludeLevel = 3, int? rightTabPos = null )
    {
      var reader = XmlReader.Create( new StringReader( string.Format( XmlTemplates.TableOfContentsXmlBase, headerStyle ?? HeaderStyle, title, rightTabPos ?? RightTabPos, BuildSwitchString( switches, lastIncludeLevel ) ) ) );
      var xml = XElement.Load( reader );
      return new TableOfContents( document, xml, headerStyle );
    }

    #endregion

    #region Private Methods

    private void InitElement( string elementName, DocX document, string headerStyle = "" )
    {
      if( elementName == "updateFields" )
      {
        if( document._settings.Descendants().Any( x => x.Name.Equals( DocX.w + elementName ) ) )
          return;

        var element = new XElement( XName.Get( elementName, DocX.w.NamespaceName ), new XAttribute( DocX.w + "val", true ) );
        document._settings.Root.Add( element );
      }
      else if( elementName == "styles" )
      {
        if( !HasStyle( document, headerStyle, "paragraph" ) )
        {
          var reader = XmlReader.Create( new StringReader( string.Format( XmlTemplates.TableOfContentsHeadingStyleBase, headerStyle ?? HeaderStyle ) ) );
          var xml = XElement.Load( reader );
          document._styles.Root.Add( xml );
        }
        if( !HasStyle( document, "TOC1", "paragraph" ) )
        {
          var reader = XmlReader.Create( new StringReader( string.Format( XmlTemplates.TableOfContentsElementStyleBase, "TOC1", "toc 1" ) ) );
          var xml = XElement.Load( reader );
          document._styles.Root.Add( xml );
        }
        if( !HasStyle( document, "TOC2", "paragraph" ) )
        {
          var reader = XmlReader.Create( new StringReader( string.Format( XmlTemplates.TableOfContentsElementStyleBase, "TOC2", "toc 2" ) ) );
          var xml = XElement.Load( reader );
          document._styles.Root.Add( xml );
        }
        if( !HasStyle( document, "TOC3", "paragraph" ) )
        {
          var reader = XmlReader.Create( new StringReader( string.Format( XmlTemplates.TableOfContentsElementStyleBase, "TOC3", "toc 3" ) ) );
          var xml = XElement.Load( reader );
          document._styles.Root.Add( xml );
        }
        if( !HasStyle( document, "TOC4", "paragraph" ) )
        {
          var reader = XmlReader.Create( new StringReader( string.Format( XmlTemplates.TableOfContentsElementStyleBase, "TOC4", "toc 4" ) ) );
          var xml = XElement.Load( reader );
          document._styles.Root.Add( xml );
        }
        if( !HasStyle( document, "Hyperlink", "character" ) )
        {
          var reader = XmlReader.Create( new StringReader( string.Format( XmlTemplates.TableOfContentsHyperLinkStyleBase ) ) );
          var xml = XElement.Load( reader );
          document._styles.Root.Add( xml );
        }
      }
    }

    private bool HasStyle( DocX document, string value, string type )
    {
      return document._styles.Descendants().Any( x => x.Name.Equals( DocX.w + "style" ) && ( x.Attribute( DocX.w + "type" ) == null || x.Attribute( DocX.w + "type" ).Value.Equals( type ) ) && x.Attribute( DocX.w + "styleId" ) != null && x.Attribute( DocX.w + "styleId" ).Value.Equals( value ) );
    }

    private static string BuildSwitchString( TableOfContentsSwitches switches, int lastIncludeLevel )
    {
      var allSwitches = Enum.GetValues( typeof( TableOfContentsSwitches ) ).Cast<TableOfContentsSwitches>();
      var switchString = "TOC";
      foreach( var s in allSwitches.Where( s => s != TableOfContentsSwitches.None && switches.HasFlag( s ) ) )
      {
        switchString += " " + s.EnumDescription();
        if( s == TableOfContentsSwitches.O )
        {
          switchString += string.Format( " '{0}-{1}'", 1, lastIncludeLevel );
        }
      }

      return switchString;
    }

    #endregion

    #region Constructor

    private TableOfContents( DocX document, XElement xml, string headerStyle )
        : base( document, xml )
    {
      InitElement( "updateFields", document );
      InitElement( "styles", document, headerStyle );
    }

    #endregion
  }
}
