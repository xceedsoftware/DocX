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
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Xceed.Document.NET
{
  public class List : InsertBeforeOrAfter
  {

    #region Private properties

    private static Random _random = new Random();

    #endregion

    #region Private Members

    private ListItemType? _listType;


    #endregion

    #region Public Properties

    public List<Paragraph> Items
    {
      get; private set;
    }

    public int NumId
    {
      get; private set;
    }

    public ListItemType? ListType
    {
      get
      {
        return _listType;
      }
    }











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
    }

    #endregion

    #region Public Methods

    public void AddItem( Paragraph paragraph )
    {
      if( paragraph.IsListItem )
      {
        var numId = paragraph.GetNumId();
        if( numId != -1 )
        {
          if( !this.CanAddListItem( paragraph ) )
            throw new InvalidOperationException( "New list items can only be added to this list if they have the same numId." );

          this.NumId = numId;
          this.SetListType( paragraph );
          this.Items.Add( paragraph );
        }
      }
    }

    public void InsertItem( Paragraph paragraph, int level, int? index = null )
    {
      if( paragraph.IsListItem )
      {
        var numId = paragraph.GetNumId();
        if( numId != -1 )
        {
          if( !this.CanAddListItem( paragraph ) )
            throw new InvalidOperationException( "New list items can only be added to this list if they have the same numId." );

          this.NumId = numId;
          this.SetListType( paragraph );

          if( index != null )
          {
            this.AddItemAtSpecificIndex( paragraph, level, index.Value );
          }
          else
          {
            this.Items.Add( paragraph );
          }
        }
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

    public void Remove()
    {
      // Remove AbstractNum and Num from numbering.xml.
      var abstractNumId = this.GetAbstractNum( this.NumId );
      if( abstractNumId != null )
      {
        abstractNumId.Remove();
      }

      var numNode = this.Document._numbering.Descendants()
                                            .Where( n => n.Name.LocalName == "num" )
                                            .FirstOrDefault( node => node.Attribute( Document.w + "numId" ).Value.Equals( this.NumId.ToString() ) );
      if( numNode != null )
      {
        numNode.Remove();
      }

      // Remove listItems from document.
      this.Items.ForEach( paragraph => paragraph.Remove( false ) );
    }

    public override Paragraph InsertParagraphBeforeSelf( string text )
    {
      return this.Items.First().InsertParagraphBeforeSelf( text, false, new Formatting() );
    }

    public override Paragraph InsertParagraphAfterSelf( string text )
    {
      return this.Items.Last().InsertParagraphAfterSelf( text, false, new Formatting() );
    }

    public override Paragraph InsertParagraphBeforeSelf( Paragraph p )
    {
      return this.Items.First().InsertParagraphBeforeSelf( p );
    }

    public override Paragraph InsertParagraphAfterSelf( Paragraph p )
    {
      return this.Items.Last().InsertParagraphBeforeSelf( p );
    }

    public override Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges )
    {
      return this.Items.First().InsertParagraphBeforeSelf( text, trackChanges, new Formatting() );
    }

    public override Paragraph InsertParagraphAfterSelf( string text, bool trackChanges )
    {
      return this.Items.Last().InsertParagraphAfterSelf( text, trackChanges, new Formatting() );
    }

    public override Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges, Formatting formatting )
    {
      return this.Items.First().InsertParagraphBeforeSelf( text, trackChanges, formatting );
    }

    public override Paragraph InsertParagraphAfterSelf( string text, bool trackChanges, Formatting formatting )
    {
      return this.Items.Last().InsertParagraphBeforeSelf( text, trackChanges, formatting );
    }

    #endregion

    #region Internal Methods

    internal void CreateNewNumberingNumId( int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool continueNumbering = false )
    {
      int numId, abstractNumId;
      XElement abstractNumTemplate;
      SetAbstractNumTemplate( listType, out numId, out abstractNumId, out abstractNumTemplate );

      // When documents contains many lists, generate different "nsid" for each abstractNumTemplate.
      // Each "nsid" value should be unique.
      var nsid = abstractNumTemplate.Element( Document.w + "nsid" );
      var val = nsid.Attribute( Document.w + "val" );
      if( val != null )
      {
        var newNSidVal = GetRandomHexNumber();
        nsid.SetAttributeValue( Document.w + "val", newNSidVal );
      }

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

    internal static string GetRandomHexNumber()
    {
      int digits = 8;
      byte[] buffer = new byte[ digits / 2 ];

      _random.NextBytes( buffer );
      string result = string.Concat( buffer.Select( x => x.ToString( "X2" ) ).ToArray() );
      if( digits % 2 == 0 )
        return result;
      return result + _random.Next( 16 ).ToString( "X" );
    }























    internal XElement GetAbstractNum( int numId )
    {
      return HelperFunctions.GetAbstractNum( this.Document, numId.ToString() );
    }

    #endregion

    #region Private Methods

    private XDocument GetListTemplate( ListItemType listType )
    {
      switch( listType )
      {
        case ListItemType.Bulleted:
          return HelperFunctions.DecompressXMLResource( HelperFunctions.GetResources( ResourceType.NumberingBullet ) );
        case ListItemType.Numbered:
          return HelperFunctions.DecompressXMLResource( HelperFunctions.GetResources( ResourceType.NumberingDecimal ) );
        default:
          throw new InvalidOperationException( string.Format( "Unable to deal with ListItemType: {0}.", listType.ToString() ) );
      }
    }

    private void SetListType( Paragraph paragraph )
    {
      if( _listType == null )
      {
        var listItemType = HelperFunctions.GetListItemType( paragraph, this.Document );
        if( listItemType != null )
        {
          _listType = listItemType.Equals( "bullet" ) ? ListItemType.Bulleted : ListItemType.Numbered;
        }
      }
    }



























































    private static int GetLevelIndex( Paragraph p )
    {
      // Extract the ilvl value from the p element.
      XElement ilvlElement = p.Xml.Descendants( XName.Get( "ilvl", Document.w.NamespaceName ) ).FirstOrDefault();
      return ilvlElement != null && ilvlElement.Attribute( XName.Get( "val", Document.w.NamespaceName ) ) != null
                         ? int.Parse( ilvlElement.Attribute( XName.Get( "val", Document.w.NamespaceName ) ).Value )
                         : -1;
    }

    private void AddItemAtSpecificIndex( Paragraph paragraph, int level, int index )
    {
      if( index >= 0 && index <= this.Items.Count )
      {
        var indexToInsertAt = List.FindIndexFromListItemsAtLevel( this.Items, level, index );

        if( indexToInsertAt != -1 )
        {
          this.Items.Insert( indexToInsertAt, paragraph );
        }
      }
    }

    private static int FindIndexFromListItemsAtLevel( List<Paragraph> items, int level, int levelIndex )
    {
      if( items != null && items.Count > 0 )
      {
        var levelItems = items.Where( p => GetLevelIndex( p ) == level ).ToList();

        if( ( levelIndex >= 0 ) && ( levelIndex < levelItems.Count ) )
        {
          var targetItem = levelItems[ levelIndex ];

          // Find the overall index of the item
          var overallIndex = items.IndexOf( items.FirstOrDefault( i => i.Xml == targetItem.Xml ) );

          return overallIndex;
        }
      }

      return -1;
    }

    private void SetAbstractNumTemplate( ListItemType listType, out int numId, out int abstractNumId, out XElement abstractNumTemplate )
    {
      ValidateDocXNumberingPartExists();
      if( Document._numbering.Root == null )
      {
        throw new InvalidOperationException( "Numbering section did not instantiate properly." );
      }

      _listType = listType;

      numId = GetMaxNumId() + 1;
      abstractNumId = GetMaxAbstractNumId() + 1;
      XDocument listTemplate = GetListTemplate( listType );

      abstractNumTemplate = listTemplate.Descendants().Single( d => d.Name.LocalName == "abstractNum" );
      abstractNumTemplate.SetAttributeValue( Document.w + "abstractNumId", abstractNumId );
    }

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
