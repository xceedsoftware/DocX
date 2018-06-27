/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Represents a List in a document.
  /// </summary>
  public class List : InsertBeforeOrAfter
  {
    #region Public Members

    /// <summary>
    /// This is a list of paragraphs that will be added to the document
    /// when the list is inserted into the document.
    /// The paragraph needs a numPr defined to be in this items collection.
    /// </summary>
    public List<Paragraph> Items
    {
      get; private set;
    }

    /// <summary>
    /// The numId used to reference the list settings in the numbering.xml
    /// </summary>
    public int NumId
    {
      get; private set;
    }

    /// <summary>
    /// The ListItemType (bullet or numbered) of the list.
    /// </summary>
    public ListItemType? ListType
    {
      get; private set;
    }

    #endregion

    #region Constructors

    internal List( DocX document, XElement xml )
        : base( document, xml )
    {
      Items = new List<Paragraph>();
      ListType = null;
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Adds an item to the list.
    /// </summary>
    /// <param name="paragraph"></param>
    /// <exception cref="InvalidOperationException">
    /// Throws an InvalidOperationException if the item cannot be added to the list.
    /// </exception>
    public void AddItem( Paragraph paragraph )
    {
      if( paragraph.IsListItem )
      {
        var numIdNode = paragraph.Xml.Descendants().First( s => s.Name.LocalName == "numId" );
        var numId = Int32.Parse( numIdNode.Attribute( DocX.w + "val" ).Value );

        if( CanAddListItem( paragraph ) )
        {
          NumId = numId;
          Items.Add( paragraph );
        }
        else
          throw new InvalidOperationException( "New list items can only be added to this list if they are have the same numId." );
      }
    }

    public void AddItemWithStartValue( Paragraph paragraph, int start )
    {
      //TODO: Update the numbering
      UpdateNumberingForLevelStartNumber( int.Parse( paragraph.IndentLevel.ToString() ), start );
      if( ContainsLevel( start ) )
        throw new InvalidOperationException( "Cannot add a paragraph with a start value if another element already exists in this list with that level." );
      AddItem( paragraph );
    }

    /// <summary>
    /// Determine if it is able to add the item to the list
    /// </summary>
    /// <param name="paragraph"></param>
    /// <returns>
    /// Return true if AddItem(...) will succeed with the given paragraph.
    /// </returns>
    public bool CanAddListItem( Paragraph paragraph )
    {
      if( paragraph.IsListItem )
      {
        //var lvlNode = paragraph.Xml.Descendants().First(s => s.Name.LocalName == "ilvl");
        var numIdNode = paragraph.Xml.Descendants().First( s => s.Name.LocalName == "numId" );
        var numId = Int32.Parse( numIdNode.Attribute( DocX.w + "val" ).Value );

        //Level = Int32.Parse(lvlNode.Attribute(DocX.w + "val").Value);
        if( NumId == 0 || ( numId == NumId && numId > 0 ) )
        {
          return true;
        }
      }
      return false;
    }

    public bool ContainsLevel( int ilvl )
    {
      return Items.Any( i => i.ParagraphNumberProperties.Descendants().First( el => el.Name.LocalName == "ilvl" ).Value == ilvl.ToString() );
    }

    #endregion

    #region Internal Methods

    internal void CreateNewNumberingNumId( int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool continueNumbering = false )
    {
      ValidateDocXNumberingPartExists();
      if( Document._numbering.Root == null )
      {
        throw new InvalidOperationException( "Numbering section did not instantiate properly." );
      }

      ListType = listType;

      var numId = GetMaxNumId() + 1;
      var abstractNumId = GetMaxAbstractNumId() + 1;

      XDocument listTemplate;
      switch( listType )
      {
        case ListItemType.Bulleted:
          listTemplate = HelperFunctions.DecompressXMLResource( "Xceed.Words.NET.Resources.numbering.default_bullet_abstract.xml.gz" );
          break;
        case ListItemType.Numbered:
          listTemplate = HelperFunctions.DecompressXMLResource( "Xceed.Words.NET.Resources.numbering.default_decimal_abstract.xml.gz" );
          break;
        default:
          throw new InvalidOperationException( string.Format( "Unable to deal with ListItemType: {0}.", listType.ToString() ) );
      }
      var abstractNumTemplate = listTemplate.Descendants().Single( d => d.Name.LocalName == "abstractNum" );
      abstractNumTemplate.SetAttributeValue( DocX.w + "abstractNumId", abstractNumId );
      var abstractNumXml = GetAbstractNumXml( abstractNumId, numId, startNumber, continueNumbering );

      var abstractNumNode = Document._numbering.Root.Descendants().LastOrDefault( xElement => xElement.Name.LocalName == "abstractNum" );
      var numXml = Document._numbering.Root.Descendants().LastOrDefault( xElement => xElement.Name.LocalName == "num" );

      if( abstractNumNode == null || numXml == null )
      {
        Document._numbering.Root.Add( abstractNumTemplate );
        Document._numbering.Root.Add( abstractNumXml );
      }
      else
      {
        abstractNumNode.AddAfterSelf( abstractNumTemplate );
        numXml.AddAfterSelf(
            abstractNumXml
        );
      }

      NumId = numId;
    }

    /// <summary>
    /// Get the abstractNum definition for the given numId
    /// </summary>
    /// <param name="numId">The numId on the pPr element</param>
    /// <returns>XElement representing the requested abstractNum</returns>
    internal XElement GetAbstractNum( int numId )
    {
      var num = Document._numbering.Descendants().First( d => d.Name.LocalName == "num" && d.GetAttribute( DocX.w + "numId" ).Equals( numId.ToString() ) );
      var abstractNumId = num.Descendants().First( d => d.Name.LocalName == "abstractNumId" ).GetAttribute( DocX.w + "val" );
      return Document._numbering.Descendants().First( d => d.Name.LocalName == "abstractNum" && d.GetAttribute( DocX.w + "abstractNumId" ).Equals( abstractNumId ) );
    }

    #endregion

    #region Private Methods

    private void UpdateNumberingForLevelStartNumber( int iLevel, int start )
    {
      var abstractNum = GetAbstractNum( NumId );
      var level = abstractNum.Descendants().First( el => el.Name.LocalName == "lvl" && el.GetAttribute( DocX.w + "ilvl" ) == iLevel.ToString() );
      level.Descendants().First( el => el.Name.LocalName == "start" ).SetAttributeValue( DocX.w + "val", start );
    }

    private XElement GetAbstractNumXml( int abstractNumId, int numId, int? startNumber, bool continueNumbering )
    {
      var start = new XElement( XName.Get( "startOverride", DocX.w.NamespaceName ), new XAttribute( DocX.w + "val", startNumber ?? 1 ) );
      var level = new XElement( XName.Get( "lvlOverride", DocX.w.NamespaceName ), new XAttribute( DocX.w + "ilvl", 0 ), start );
      var element = new XElement( XName.Get( "abstractNumId", DocX.w.NamespaceName ), new XAttribute( DocX.w + "val", abstractNumId ) );

      return continueNumbering
          ? new XElement( XName.Get( "num", DocX.w.NamespaceName ), new XAttribute( DocX.w + "numId", numId ), element )
          : new XElement( XName.Get( "num", DocX.w.NamespaceName ), new XAttribute( DocX.w + "numId", numId ), element, level );
    }

    /// <summary>
    /// Method to determine the last numId for a list element. 
    /// Also useful for determining the next numId to use for inserting a new list element into the document.
    /// </summary>
    /// <returns>
    /// 0 if there are no elements in the list already.
    /// Increment the return for the next valid value of a new list element.
    /// </returns>
    private int GetMaxNumId()
    {
      const int defaultValue = 0;
      if( Document._numbering == null )
        return defaultValue;

      var numlist = Document._numbering.Descendants().Where( d => d.Name.LocalName == "num" ).ToList();
      if( numlist.Any() )
        return numlist.Attributes( DocX.w + "numId" ).Max( e => int.Parse( e.Value ) );
      return defaultValue;
    }

    /// <summary>
    /// Method to determine the last abstractNumId for a list element.
    /// Also useful for determining the next abstractNumId to use for inserting a new list element into the document.
    /// </summary>
    /// <returns>
    /// -1 if there are no elements in the list already.
    /// Increment the return for the next valid value of a new list element.
    /// </returns>
    private int GetMaxAbstractNumId()
    {
      const int defaultValue = -1;

      if( Document._numbering == null )
        return defaultValue;

      var numlist = Document._numbering.Descendants().Where( d => d.Name.LocalName == "abstractNum" ).ToList();
      if( numlist.Any() )
      {
        var maxAbstractNumId = numlist.Attributes( DocX.w + "abstractNumId" ).Max( e => int.Parse( e.Value ) );
        return maxAbstractNumId;
      }
      return defaultValue;
    }

    private void ValidateDocXNumberingPartExists()
    {
      var numberingUri = new Uri( "/word/numbering.xml", UriKind.Relative );

      // If the internal document contains no /word/numbering.xml create one.
      if( !Document._package.PartExists( numberingUri ) )
        Document._numbering = HelperFunctions.AddDefaultNumberingXml( Document._package );
    }

    #endregion
  }
}
