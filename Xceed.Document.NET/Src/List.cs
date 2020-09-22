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
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;

namespace Xceed.Document.NET
{
  /// <summary>
  /// Represents a List in a document.
  /// </summary>
  public class List : InsertBeforeOrAfter
  {
    #region Public Properties

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

    #region Internal Properties

    public int AbstractNumId
    {
      get
      {
        var abstractNumIdString = HelperFunctions.GetAbstractNumIdValue( this.Document, this.NumId.ToString() );
        return ( abstractNumIdString != null ) ? Int32.Parse( abstractNumIdString ) : 0;
      }
    }

    #endregion

    #region Constructors

    internal List( Document document, XElement xml )
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
        var numId = paragraph.GetNumId();
        if( numId == -1 )
          return;

        if( this.CanAddListItem( paragraph ) )
        {
          this.NumId = numId;
          if( this.ListType == null )
          {
            var listItemType = HelperFunctions.GetListItemType( paragraph, this.Document );
            if( listItemType != null )
            {
              this.ListType = listItemType.Equals( "bullet" ) ? ListItemType.Bulleted : ListItemType.Numbered;
            }
          }
          this.Items.Add( paragraph );
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
        var numId = paragraph.GetNumId();
        if( numId == -1 )
          return false;

        if( ( this.NumId == 0 )
          || ( ( numId == this.NumId ) && ( numId > 0 ) ) )
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
          listTemplate = HelperFunctions.DecompressXMLResource( HelperFunctions.GetResources( ResourceType.NumberingBullet ) );
          break;
        case ListItemType.Numbered:
          listTemplate = HelperFunctions.DecompressXMLResource( HelperFunctions.GetResources( ResourceType.NumberingDecimal ) );
          break;
        default:
          throw new InvalidOperationException( string.Format( "Unable to deal with ListItemType: {0}.", listType.ToString() ) );
      }
      var abstractNumTemplate = listTemplate.Descendants().Single( d => d.Name.LocalName == "abstractNum" );
      abstractNumTemplate.SetAttributeValue( Document.w + "abstractNumId", abstractNumId );
      var abstractNumXml = GetAbstractNumXml( abstractNumId, numId, startNumber, level, continueNumbering );

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
      return HelperFunctions.GetAbstractNum( this.Document, numId.ToString() );
    }

    #endregion

    #region Private Methods

    private void UpdateNumberingForLevelStartNumber( int iLevel, int start )
    {
      // Find num node in numbering.
      var documentNumberingDescendants = this.Document._numbering.Descendants();
      var numNodes = documentNumberingDescendants.Where( n => n.Name.LocalName == "num" );
      var numNode = numNodes.FirstOrDefault( node => node.Attribute( Document.w + "numId" ).Value.Equals( this.NumId.ToString() ) );
      if( numNode != null )
      {
        var isStartOverrideUpdated = false;
        var lvlOverrides = numNode.Elements( XName.Get( "lvlOverride", Document.w.NamespaceName ) );
        if( lvlOverrides.Count() > 0 )
        {
          foreach( var singleLvlOverride in lvlOverrides )
          {
            var ilvl = singleLvlOverride.GetAttribute( XName.Get( "ilvl", Document.w.NamespaceName ) );
            // Found same level, update its startOverride.
            if( !string.IsNullOrEmpty( ilvl ) && ( ilvl == iLevel.ToString() ) )
            {
              var startOverride = singleLvlOverride.Element( XName.Get( "startOverride", Document.w.NamespaceName ) );
              if( startOverride != null )
              {
                startOverride.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), start );
              }
              else
              {
                singleLvlOverride.Add( new XElement( XName.Get( "startOverride", Document.w.NamespaceName ), new XAttribute( Document.w + "val", start ) ) );
              }

              isStartOverrideUpdated = true;
              break;
            }
          }
        }

        if( !isStartOverrideUpdated )
        {
          var startOverride = new XElement( XName.Get( "startOverride", Document.w.NamespaceName ), new XAttribute( Document.w + "val", start ) );
          var levelOverride = new XElement( XName.Get( "lvlOverride", Document.w.NamespaceName ), new XAttribute( Document.w + "ilvl", iLevel ), startOverride );
          numNode.Add( levelOverride );
        }
      }
    }

    private XElement GetAbstractNumXml( int abstractNumId, int numId, int? startNumber, int level, bool continueNumbering )
    {
      var start = new XElement( XName.Get( "startOverride", Document.w.NamespaceName ), new XAttribute( Document.w + "val", startNumber ?? 1 ) );
      var levelOverride = new XElement( XName.Get( "lvlOverride", Document.w.NamespaceName ), new XAttribute( Document.w + "ilvl", level ), start );
      var element = new XElement( XName.Get( "abstractNumId", Document.w.NamespaceName ), new XAttribute( Document.w + "val", abstractNumId ) );

      return continueNumbering
          ? new XElement( XName.Get( "num", Document.w.NamespaceName ), new XAttribute( Document.w + "numId", numId ), element )
          : new XElement( XName.Get( "num", Document.w.NamespaceName ), new XAttribute( Document.w + "numId", numId ), element, levelOverride );
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
        return numlist.Attributes( Document.w + "numId" ).Max( e => int.Parse( e.Value ) );
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
        var maxAbstractNumId = numlist.Attributes( Document.w + "abstractNumId" ).Max( e => int.Parse( e.Value ) );
        return maxAbstractNumId;
      }
      return defaultValue;
    }

    private void ValidateDocXNumberingPartExists()
    {
      var numberingUri = new Uri( "/word/numbering.xml", UriKind.Relative );

      // If the internal document contains no /word/numbering.xml create one.
      if( !Document._package.PartExists( numberingUri ) )
      {
        Document._numbering = HelperFunctions.AddDefaultNumberingXml( Document._package );
        Document._numberingPart = Document._package.GetPart( numberingUri );
      }
    }

    #endregion
  }
}
