/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2026 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Globalization;
using System.Collections.ObjectModel;

namespace Xceed.Document.NET
{
  public class Row : Container
  {
    #region Internal Members

    internal Table _table;
    internal string _paraId;

    #endregion

    #region Public Properties

    public Int32 ColumnCount
    {
      get
      {
        int gridSpanSum = this.GridAfter;

        // Foreach each Cell between startIndex and endIndex inclusive.
        foreach( Cell c in Cells )
        {
          if( c.GridSpan != 0 )
          {
            gridSpanSum += ( c.GridSpan - 1 );
          }
        }

        // return cells count + count of spanned cells
        return Cells.Count + gridSpanSum;
      }
    }

    public int GridAfter
    {
      get
      {
        var trPr = this.Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
        if( trPr != null )
        {
          var gridAfter = trPr.Element( XName.Get( "gridAfter", Document.w.NamespaceName ) );
          var gridAfterAttrVal = gridAfter?.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
          if( gridAfterAttrVal != null )
          {
            return int.Parse( gridAfterAttrVal.Value );
          }
        }

        return 0;
      }

      internal set
      {
        if( value < 0 )
          throw new InvalidDataException( "GridAfter value must be greater than 0" );

        var trPr = this.Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
        if( trPr == null )
        {
          this.Xml.AddFirst( new XElement( XName.Get( "trPr", Document.w.NamespaceName ), string.Empty ) );
          trPr = this.Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
        }

        var gridAfter = trPr.Element( XName.Get( "gridAfter", Document.w.NamespaceName ) );

        if( ( gridAfter == null ) && ( value > 0 ) )
        {
          trPr.Add( new XElement( XName.Get( "gridAfter", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), value ) ) );
        }

        if( ( gridAfter != null ) && ( value == 0 ) )
        {
          gridAfter.Remove();
        }
      }
    }

    public List<Cell> Cells
    {
      get
      {
        List<Cell> cells =
        (
            from c in this.Xml.Descendants( XName.Get( "tc", Document.w.NamespaceName ) )
            where ( this.GetParentRow( c ) == this.Xml )
            select new Cell( this, this.Document, c )
        ).ToList();
        return cells;
      }
    }

    public override ReadOnlyCollection<Paragraph> Paragraphs
    {
      get
      {
        var paragraphs =
        (
            from p in Xml.Descendants( Document.w + "p" )
            select new Paragraph( Document, p, 0 )
        ).ToList();

        foreach( Paragraph p in paragraphs )
        {
          p.PackagePart = _table.PackagePart;
        }

        return paragraphs.AsReadOnly();
      }
    }

    public double Height
    {
      get
      {
        /*
        * Get the trPr (table row properties) element for this Row,
        * null will be return if no such element exists.
        */
        XElement trPr = Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );

        // If trPr is null, this row contains no height information.
        // Get the trHeight element for this Row,
        // null will be return if no such element exists.
        XElement trHeight = trPr?.Element( XName.Get( "trHeight", Document.w.NamespaceName ) );

        // If trHeight is null, this row contains no height information.
        // Get the val attribute for this trHeight element.
        XAttribute val = trHeight?.Attribute( XName.Get( "val", Document.w.NamespaceName ) );

        // If w is null, this cell contains no width information.
        if( val == null )
          return double.NaN;

        // If val is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
        double heightInWordUnits;
        if( !HelperFunctions.TryParseDouble( val.Value, out heightInWordUnits ) )
        {
          val.Remove();
          return double.NaN;
        }

        // Using 20 to match Document._pageSizeMultiplier.
        return ( heightInWordUnits / 20 );
      }

      set
      {
        SetHeight( value, true );
      }
    }

    public double MinHeight
    {
      get
      {
        return Height;
      }
      set
      {
        SetHeight( value, false );
      }
    }

    public bool TableHeader
    {
      get
      {
        XElement trPr = Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
        XElement tblHeader = trPr?.Element( XName.Get( "tblHeader", Document.w.NamespaceName ) );
        return tblHeader != null;
      }
      set
      {
        XElement trPr = Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
        if( trPr == null )
        {
          Xml.SetElementValue( XName.Get( "trPr", Document.w.NamespaceName ), string.Empty );
          trPr = Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
        }

        XElement tblHeader = trPr.Element( XName.Get( "tblHeader", Document.w.NamespaceName ) );

        if( tblHeader == null && value )
          trPr.SetElementValue( XName.Get( "tblHeader", Document.w.NamespaceName ), string.Empty );

        if( tblHeader != null && !value )
          tblHeader.Remove();
      }
    }

    public bool BreakAcrossPages
    {
      get
      {
        var trPr = Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
        var cantSplit = trPr?.Element( XName.Get( "cantSplit", Document.w.NamespaceName ) );
        return cantSplit == null;
      }
      set
      {
        var trPrXName = XName.Get( "trPr", Document.w.NamespaceName );
        var cantSplitXName = XName.Get( "cantSplit", Document.w.NamespaceName );

        if( value )
        {
          var trPr = Xml.Element( trPrXName );
          var cantSplit = trPr?.Element( cantSplitXName );
          if( cantSplit != null )
            cantSplit.Remove();
        }
        else
        {
          var trPr = Xml.Element( trPrXName );
          if( trPr == null )
          {
            Xml.SetElementValue( trPrXName, string.Empty );
            trPr = Xml.Element( trPrXName );
          }
          var cantSplit = trPr.Element( cantSplitXName );
          if( cantSplit == null )
          {
            trPr.SetElementValue( cantSplitXName, string.Empty );
          }
        }
      }
    }

    #endregion

    #region Constructors

    internal Row( Table table, Document document, XElement xml )
        : base( document, xml )
    {
      _table = table;
      this.PackagePart = table.PackagePart;
    }



    #endregion

    #region Public Methods

    public void Remove()
    {
      XElement table = Xml.Parent;

      Xml.Remove();
      if( !table.Elements( XName.Get( "tr", Document.w.NamespaceName ) ).Any() )
        table.Remove();
    }

    public void MergeCells( int startIndex, int endIndex )
    {
      // Check for valid start and end indexes.
      if( startIndex < 0 || endIndex <= startIndex || endIndex > Cells.Count + 1 )
        throw new IndexOutOfRangeException();

      // The sum of all merged gridSpans.
      int gridSpanSum = 0;

      // Foreach each Cell between startIndex and endIndex inclusive.
      var cellsToMerge = this.Cells.Where( ( z, i ) => i > startIndex && i <= endIndex );
      foreach( var c in cellsToMerge )
      {
        var tcPr = c.Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        var gridSpan = tcPr?.Element( XName.Get( "gridSpan", Document.w.NamespaceName ) );
        if( gridSpan != null )
        {
          var val = gridSpan.Attribute( XName.Get( "val", Document.w.NamespaceName ) );

          int value;
          if( val != null && HelperFunctions.TryParseInt( val.Value, out value ) )
            gridSpanSum += value - 1;
        }

        // Add this cells Pragraph to the merge start Cell.
        this.Cells[ startIndex ].Xml.Add( c.Xml.Elements( XName.Get( "p", Document.w.NamespaceName ) ) );

        this.Cells[ startIndex ].Width += c.Width;

        // Remove this Cell.
        c.Xml.Remove();
      }

      // Trim cell's paragraphs to remove extra blank lines, if any
      int index = 0;
      do
      {
        var cellsStartIndexParagraphs = Cells[ startIndex ].Paragraphs;
        // If the cell doesn't have multiple paragraphs, leave the loop
        if( cellsStartIndexParagraphs.Count < 2 )
          break;

        // Remove the last paragraph if it's a blank line, otherwise trimming is done
        index = cellsStartIndexParagraphs.Count - 1;
        if( cellsStartIndexParagraphs[ index ].Text.Trim() == "" )
          cellsStartIndexParagraphs[ index ].Remove( false );
        else
          break;
      } while( true );

      /* 
       * Get the tcPr (table cell properties) element for the first cell in this merge,
       * null will be returned if no such element exists.
       */
      XElement start_tcPr = Cells[ startIndex ].Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      if( start_tcPr == null )
      {
        Cells[ startIndex ].Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
        start_tcPr = Cells[ startIndex ].Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      }

      /* 
       * Get the gridSpan element of this row,
       * null will be returned if no such element exists.
       */
      XElement start_gridSpan = start_tcPr.Element( XName.Get( "gridSpan", Document.w.NamespaceName ) );
      if( start_gridSpan == null )
      {
        start_tcPr.SetElementValue( XName.Get( "gridSpan", Document.w.NamespaceName ), string.Empty );
        start_gridSpan = start_tcPr.Element( XName.Get( "gridSpan", Document.w.NamespaceName ) );
      }

      /* 
       * Get the val attribute of this row,
       * null will be returned if no such element exists.
       */
      XAttribute start_val = start_gridSpan.Attribute( XName.Get( "val", Document.w.NamespaceName ) );

      int start_value = 0;
      if( start_val != null )
        if( HelperFunctions.TryParseInt( start_val.Value, out start_value ) )
          gridSpanSum += start_value - 1;

      // Set the val attribute to the number of merged cells.
      start_gridSpan.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), ( gridSpanSum + ( endIndex - startIndex + 1 ) ).ToString() );
    }

    #endregion

    #region Internal Methods





    #endregion

    #region Private Methods

    private void SetHeight( double height, bool isHeightExact )
    {
      XElement trPr = Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
      if( trPr == null )
      {
        Xml.SetElementValue( XName.Get( "trPr", Document.w.NamespaceName ), string.Empty );
        trPr = Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
      }

      XElement tc = Xml.Element( XName.Get( "tc", Document.w.NamespaceName ) );
      if( tc != null )
      {
        trPr.Remove();
        tc.AddBeforeSelf( trPr );
      }

      XElement trHeight = trPr.Element( XName.Get( "trHeight", Document.w.NamespaceName ) );
      if( trHeight == null )
      {
        trPr.SetElementValue( XName.Get( "trHeight", Document.w.NamespaceName ), string.Empty );
        trHeight = trPr.Element( XName.Get( "trHeight", Document.w.NamespaceName ) );
      }

      // The hRule attribute needs to be set.
      trHeight.SetAttributeValue( XName.Get( "hRule", Document.w.NamespaceName ), isHeightExact ? "exact" : "atLeast" );

      // Using 20 to match Document._pageSizeMultiplier.
      trHeight.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), ( (int)( Math.Round( height * 20, 0 ) ) ).ToString( CultureInfo.InvariantCulture ) );
    }

    private XElement GetParentRow( XElement xElement )
    {
      while( ( xElement != null ) && ( xElement.Name != XName.Get( "tr", Document.w.NamespaceName ) ) )
      {
        xElement = xElement.Parent;
      }

      return xElement;
    }

    #endregion
  }
}
