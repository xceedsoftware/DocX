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
using System.Globalization;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace Xceed.Document.NET
{
  public class Cell : Container
  {
    #region Static variables

    private static readonly int DefaultWidth = 2000;

    #endregion

    #region Private Members

    private string _paraId;

    #endregion

    #region Internal Members

    internal Row _row;
    internal ShadingPattern _shadingPattern;

    #endregion

    #region Public Properties

    public override ReadOnlyCollection<Paragraph> Paragraphs
    {
      get
      {
        var paragraphs = base.Paragraphs;

        foreach( Paragraph p in paragraphs )
        {
          p.PackagePart = _row._table.PackagePart;
        }

        return paragraphs;
      }
    }

    // <summary>
    // Gets or Sets this Cells vertical alignment.
    // </summary>
    // <example>
    // Creates a table with 3 cells and sets the vertical alignment of each to 1 of the 3 available options.
    // <code>
    // Create a new document.
    //using(var document = DocX.Create("Test.docx"))
    //{
    //    // Insert a Table into this document.
    //    Table t = document.InsertTable(3, 1);
    //
    //    // Set the design of the Table such that we can easily identify cell boundaries.
    //    t.Design = TableDesign.TableGrid;
    //
    //    // Set the height of the row bigger than default.
    //    // We need to be able to see the difference in vertical cell alignment options.
    //    t.Rows[0].Height = 100;
    //
    //    // Set the vertical alignment of cell0 to top.
    //    Cell c0 = t.Rows[0].Cells[0];
    //    c0.InsertParagraph("VerticalAlignment.Top");
    //    c0.VerticalAlignment = VerticalAlignment.Top;
    //
    //    // Set the vertical alignment of cell1 to center.
    //    Cell c1 = t.Rows[0].Cells[1];
    //    c1.InsertParagraph("VerticalAlignment.Center");
    //    c1.VerticalAlignment = VerticalAlignment.Center;
    //
    //    // Set the vertical alignment of cell2 to bottom.
    //    Cell c2 = t.Rows[0].Cells[2];
    //    c2.InsertParagraph("VerticalAlignment.Bottom");
    //    c2.VerticalAlignment = VerticalAlignment.Bottom;
    //
    //    // Save the document.
    //    document.Save();
    //}
    // </code>
    // </example>
    public VerticalAlignment VerticalAlignment
    {
      get
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );

        // If tcPr is null, this cell contains no width information.
        // Get the vAlign (table cell vertical alignment) element for this Cell,
        // null will be return if no such element exists.
        XElement vAlign = tcPr?.Element( XName.Get( "vAlign", Document.w.NamespaceName ) );

        // If vAlign is null, this cell contains no vertical alignment information.
        // Get the val attribute of the vAlign element.
        XAttribute val = vAlign?.Attribute( XName.Get( "val", Document.w.NamespaceName ) );

        // If val is null, this cell contains no vAlign information.
        if( val == null )
          return VerticalAlignment.Top;

        // If val is not a VerticalAlign enum, something is wrong with this attributes value, so remove it and return VerticalAlignment.Center;
        try
        {
          return ( VerticalAlignment )Enum.Parse( typeof( VerticalAlignment ), val.Value, true );
        }

        catch
        {
          val.Remove();
          return VerticalAlignment.Top;
        }
      }

      set
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        /*
         * Get the vAlign (table cell vertical alignment) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement vAlign = tcPr.Element( XName.Get( "vAlign", Document.w.NamespaceName ) );
        if( vAlign == null )
        {
          tcPr.SetElementValue( XName.Get( "vAlign", Document.w.NamespaceName ), string.Empty );
          vAlign = tcPr.Element( XName.Get( "vAlign", Document.w.NamespaceName ) );
        }

        // Set the VerticalAlignment in 'val'
        vAlign.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), value.ToString().ToLower() );
      }
    }

    [Obsolete( "This property is obsolete and should no longer be used. Use the ShadingPattern property instead." )]
    public Xceed.Drawing.Color Shading
    {
      get
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );

        // If tcPr is null, this cell contains no Color information.
        // Get the shd (table shade) element for this Cell,
        // null will be return if no such element exists.
        XElement shd = tcPr?.Element( XName.Get( "shd", Document.w.NamespaceName ) );

        // If shd is null, this cell contains no Color information.
        // Get the w attribute of the tcW element.
        XAttribute fill = shd?.Attribute( XName.Get( "fill", Document.w.NamespaceName ) );

        // If fill is null, this cell contains no Color information.
        if( fill == null )
          return Xceed.Drawing.Color.White;

        return HelperFunctions.GetColorFromHtml( fill.Value );
      }

      set
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        /*
         * Get the shd (table shade) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement shd = tcPr.Element( XName.Get( "shd", Document.w.NamespaceName ) );
        if( shd == null )
        {
          tcPr.SetElementValue( XName.Get( "shd", Document.w.NamespaceName ), string.Empty );
          shd = tcPr.Element( XName.Get( "shd", Document.w.NamespaceName ) );
        }

        // The val attribute needs to be set to clear
        shd.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), "clear" );

        // The color attribute needs to be set to auto
        shd.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), "auto" );

        // The fill attribute needs to be set to the hex for this Color.
        shd.SetAttributeValue( XName.Get( "fill", Document.w.NamespaceName ), value.ToHex() );
      }
    }

    #region ShadingPattern

    public ShadingPattern ShadingPattern
    {
      get
      {
        if( _shadingPattern != null )
          return _shadingPattern;

        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );

        // If tcPr is null, this cell contains no Pattern information.
        // Get the shd (table shade) element for this Cell,
        // null will be return if no such element exists.
        XElement shd = tcPr?.Element( XName.Get( "shd", Document.w.NamespaceName ) );


        _shadingPattern = new ShadingPattern();

        // If shd is not null, this cell contains Pattern information, else return an empty ShadingPattern.
        if( shd != null )
        {
          // Get the w attribute of the tcW element.
          XAttribute fill = shd.Attribute( XName.Get( "fill", Document.w.NamespaceName ) );
          XAttribute style = shd.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
          XAttribute styleColor = shd.Attribute( XName.Get( "color", Document.w.NamespaceName ) );

          _shadingPattern.Fill = HelperFunctions.GetColorFromHtml( fill.Value );
          _shadingPattern.Style = HelperFunctions.GetTablePatternStyleFromValue( style.Value );
          _shadingPattern.StyleColor = HelperFunctions.GetColorFromHtml( styleColor.Value );
        }

        _shadingPattern.PropertyChanged += this.ShadingPattern_PropertyChanged;

        return _shadingPattern;
      }

      set
      {
        if( _shadingPattern != null )
        {
          _shadingPattern.PropertyChanged -= this.ShadingPattern_PropertyChanged;
        }

        _shadingPattern = value;

        if( _shadingPattern != null )
        {
          _shadingPattern.PropertyChanged += this.ShadingPattern_PropertyChanged;
        }

        this.UpdateShadingPatternXml();
      }
    }

    #endregion  //ShadingPattern

    public double Width
    {
      get
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );

        // If tcPr is null, this cell contains no width information.
        // Get the tcW (table cell width) element for this Cell,
        // null will be return if no such element exists.
        XElement tcW = tcPr?.Element( XName.Get( "tcW", Document.w.NamespaceName ) );

        // If tcW is null, this cell contains no width information.
        // Get the w attribute of the tcW element.
        XAttribute w = tcW?.Attribute( XName.Get( "w", Document.w.NamespaceName ) );
        XAttribute type = tcW?.Attribute( XName.Get( "type", Document.w.NamespaceName ) );

        // If w is null, this cell contains no width information.
        if( w == null )
          return double.NaN;

        // If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
        double widthInWordUnits;
        if( !HelperFunctions.TryParseDouble( w.Value, out widthInWordUnits ) )
        {
          w.Remove();
          return double.NaN;
        }

        if( type != null )
        {
          if( type.Value == "pct" )
          {
            if( this._row._table.ColumnWidths != null )
              return ( ( widthInWordUnits / 5000d ) * this._row._table.ColumnWidths.Sum() );
          }
          else if( type.Value == "auto" )
          {
            var cellIndex = this._row.Cells.FindIndex( x => x.Xml == this.Xml );
            if( ( cellIndex >= 0 ) && ( this._row._table.ColumnWidths != null ) )
              return this._row._table.ColumnWidths[ cellIndex ];
          }
        }
        // Using 20 to match Document._pageSizeMultiplier.
        return ( widthInWordUnits / 20 );
      }

      set
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        /*
         * Get the tcW (table cell width) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcW = tcPr.Element( XName.Get( "tcW", Document.w.NamespaceName ) );
        if( tcW == null )
        {
          tcPr.SetElementValue( XName.Get( "tcW", Document.w.NamespaceName ), string.Empty );
          tcW = tcPr.Element( XName.Get( "tcW", Document.w.NamespaceName ) );
        }

        if( value == -1 )
        {
          // remove cell width; due to set on table prop.
          tcW.Remove();
          return;
        }

        // The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
        tcW.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), "dxa" );

        // Using 20 to match Document._pageSizeMultiplier.
        tcW.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), ( value * 20 ).ToString( CultureInfo.InvariantCulture ) );
      }
    }

    public double MarginLeft
    {
      get
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );

        // If tcPr is null, this cell contains no width information.
        if( tcPr == null )
          return double.NaN;

        /*
         * Get the tcMar
         * 
         */
        XElement tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );

        // If tcMar is null, this cell contains no margin information.
        // Get the left (LeftMargin) element
        XElement tcMarLeft = tcMar?.Element( XName.Get( "left", Document.w.NamespaceName ) );

        // If tcMarLeft is null, this cell contains no left margin information.
        // Get the w attribute of the tcMarLeft element.
        XAttribute w = tcMarLeft?.Attribute( XName.Get( "w", Document.w.NamespaceName ) );

        // If w is null, this cell contains no width information.
        if( w == null )
          return double.NaN;

        // If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
        double leftMarginInWordUnits;
        if( !HelperFunctions.TryParseDouble( w.Value, out leftMarginInWordUnits ) )
        {
          w.Remove();
          return double.NaN;
        }

        // Using 20 to match Document._pageSizeMultiplier.
        return ( leftMarginInWordUnits / 20 );
      }

      set
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        /*
         * Get the tcMar (table cell margin) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );
        if( tcMar == null )
        {
          tcPr.SetElementValue( XName.Get( "tcMar", Document.w.NamespaceName ), string.Empty );
          tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );
        }

        /*
         * Get the left (table cell left margin) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcMarLeft = tcMar.Element( XName.Get( "left", Document.w.NamespaceName ) );
        if( tcMarLeft == null )
        {
          tcMar.SetElementValue( XName.Get( "left", Document.w.NamespaceName ), string.Empty );
          tcMarLeft = tcMar.Element( XName.Get( "left", Document.w.NamespaceName ) );
        }

        // The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
        tcMarLeft.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), "dxa" );

        // Using 20 to match Document._pageSizeMultiplier.
        tcMarLeft.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), ( value * 20 ).ToString( CultureInfo.InvariantCulture ) );
      }
    }

    public double MarginRight
    {
      get
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );

        // If tcPr is null, this cell contains no width information.
        if( tcPr == null )
          return double.NaN;

        /*
         * Get the tcMar
         * 
         */
        XElement tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );

        // If tcMar is null, this cell contains no margin information.
        // Get the right (RightMargin) element
        XElement tcMarRight = tcMar?.Element( XName.Get( "right", Document.w.NamespaceName ) );

        // If tcMarRight is null, this cell contains no right margin information.
        // Get the w attribute of the tcMarRight element.
        XAttribute w = tcMarRight?.Attribute( XName.Get( "w", Document.w.NamespaceName ) );

        // If w is null, this cell contains no width information.
        if( w == null )
          return double.NaN;

        // If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
        double rightMarginInWordUnits;
        if( !HelperFunctions.TryParseDouble( w.Value, out rightMarginInWordUnits ) )
        {
          w.Remove();
          return double.NaN;
        }

        // Using 20 to match Document._pageSizeMultiplier.
        return ( rightMarginInWordUnits / 20 );
      }

      set
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        /*
         * Get the tcMar (table cell margin) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );
        if( tcMar == null )
        {
          tcPr.SetElementValue( XName.Get( "tcMar", Document.w.NamespaceName ), string.Empty );
          tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );
        }

        /*
         * Get the right (table cell right margin) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcMarRight = tcMar.Element( XName.Get( "right", Document.w.NamespaceName ) );
        if( tcMarRight == null )
        {
          tcMar.SetElementValue( XName.Get( "right", Document.w.NamespaceName ), string.Empty );
          tcMarRight = tcMar.Element( XName.Get( "right", Document.w.NamespaceName ) );
        }

        // The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
        tcMarRight.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), "dxa" );

        // Using 20 to match Document._pageSizeMultiplier.
        tcMarRight.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), ( value * 20 ).ToString( CultureInfo.InvariantCulture ) );
      }
    }

    public double MarginTop
    {
      get
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );

        // If tcPr is null, this cell contains no width information.
        if( tcPr == null )
          return double.NaN;

        /*
         * Get the tcMar
         * 
         */
        XElement tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );

        // If tcMar is null, this cell contains no margin information.
        // Get the top (TopMargin) element
        XElement tcMarTop = tcMar?.Element( XName.Get( "top", Document.w.NamespaceName ) );

        // If tcMarTop is null, this cell contains no top margin information.
        // Get the w attribute of the tcMarTop element.
        XAttribute w = tcMarTop?.Attribute( XName.Get( "w", Document.w.NamespaceName ) );

        // If w is null, this cell contains no width information.
        if( w == null )
          return double.NaN;

        // If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
        double topMarginInWordUnits;
        if( !HelperFunctions.TryParseDouble( w.Value, out topMarginInWordUnits ) )
        {
          w.Remove();
          return double.NaN;
        }

        // Using 20 to match Document._pageSizeMultiplier.
        return ( topMarginInWordUnits / 20 );
      }

      set
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        /*
         * Get the tcMar (table cell margin) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );
        if( tcMar == null )
        {
          tcPr.SetElementValue( XName.Get( "tcMar", Document.w.NamespaceName ), string.Empty );
          tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );
        }

        /*
         * Get the top (table cell top margin) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcMarTop = tcMar.Element( XName.Get( "top", Document.w.NamespaceName ) );
        if( tcMarTop == null )
        {
          tcMar.SetElementValue( XName.Get( "top", Document.w.NamespaceName ), string.Empty );
          tcMarTop = tcMar.Element( XName.Get( "top", Document.w.NamespaceName ) );
        }

        // The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
        tcMarTop.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), "dxa" );

        // Using 20 to match Document._pageSizeMultiplier.
        tcMarTop.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), ( value * 20 ).ToString( CultureInfo.InvariantCulture ) );
      }
    }

    public double MarginBottom
    {
      get
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );

        // If tcPr is null, this cell contains no width information.
        if( tcPr == null )
          return double.NaN;

        /*
         * Get the tcMar
         * 
         */
        XElement tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );

        // If tcMar is null, this cell contains no margin information.
        // Get the bottom (BottomMargin) element
        XElement tcMarBottom = tcMar?.Element( XName.Get( "bottom", Document.w.NamespaceName ) );

        // If tcMarBottom is null, this cell contains no bottom margin information.
        // Get the w attribute of the tcMarBottom element.
        XAttribute w = tcMarBottom?.Attribute( XName.Get( "w", Document.w.NamespaceName ) );

        // If w is null, this cell contains no width information.
        if( w == null )
          return double.NaN;

        // If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
        double bottomMarginInWordUnits;
        if( !HelperFunctions.TryParseDouble( w.Value, out bottomMarginInWordUnits ) )
        {
          w.Remove();
          return double.NaN;
        }

        // Using 20 to match Document._pageSizeMultiplier.
        return ( bottomMarginInWordUnits / 20 );
      }

      set
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        /*
         * Get the tcMar (table cell margin) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );
        if( tcMar == null )
        {
          tcPr.SetElementValue( XName.Get( "tcMar", Document.w.NamespaceName ), string.Empty );
          tcMar = tcPr.Element( XName.Get( "tcMar", Document.w.NamespaceName ) );
        }

        /*
         * Get the bottom (table cell bottom margin) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcMarBottom = tcMar.Element( XName.Get( "bottom", Document.w.NamespaceName ) );
        if( tcMarBottom == null )
        {
          tcMar.SetElementValue( XName.Get( "bottom", Document.w.NamespaceName ), string.Empty );
          tcMarBottom = tcMar.Element( XName.Get( "bottom", Document.w.NamespaceName ) );
        }

        // The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
        tcMarBottom.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), "dxa" );

        // Using 20 to match Document._pageSizeMultiplier.
        tcMarBottom.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), ( value * 20 ).ToString( CultureInfo.InvariantCulture ) );
      }
    }

    public Xceed.Drawing.Color FillColor
    {
      get
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        XElement shd = tcPr?.Element( XName.Get( "shd", Document.w.NamespaceName ) );
        XAttribute fill = shd?.Attribute( XName.Get( "fill", Document.w.NamespaceName ) );
        if( fill == null )
          return Xceed.Drawing.Color.Transparent;

        return Xceed.Drawing.Color.Parse( fill.Value.Replace( "#", "" ) );
      }

      set
      {
        /*
         * Get the tcPr (table cell properties) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        /*
         * Get the tcW (table cell width) element for this Cell,
         * null will be return if no such element exists.
         */
        XElement shd = tcPr.Element( XName.Get( "shd", Document.w.NamespaceName ) );
        if( shd == null )
        {
          tcPr.SetElementValue( XName.Get( "shd", Document.w.NamespaceName ), string.Empty );
          shd = tcPr.Element( XName.Get( "shd", Document.w.NamespaceName ) );
        }

        shd.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), "clear" );
        shd.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), "auto" );
        shd.SetAttributeValue( XName.Get( "fill", Document.w.NamespaceName ), value.ToHex() );
      }
    }

    public TextDirection TextDirection
    {
      get
      {
        var tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        var textDirection = tcPr?.Element( XName.Get( "textDirection", Document.w.NamespaceName ) );
        var val = textDirection?.Attribute( XName.Get( "val", Document.w.NamespaceName ) );

        if( val != null )
        {
          TextDirection textDirectionValue;
          var success = Enum.TryParse( val.Value, out textDirectionValue );
          if( success )
            return textDirectionValue;
          else
          {
            val.Remove();
          }
        }

        return TextDirection.right;
      }

      set
      {
        var tcPrXName = XName.Get( "tcPr", Document.w.NamespaceName );
        var textDirectionXName = XName.Get( "textDirection", Document.w.NamespaceName );

        var tcPr = Xml.Element( tcPrXName );
        if( tcPr == null )
        {
          Xml.SetElementValue( tcPrXName, string.Empty );
          tcPr = Xml.Element( tcPrXName );
        }

        var textDirection = tcPr.Element( textDirectionXName );
        if( textDirection == null )
        {
          tcPr.SetElementValue( textDirectionXName, string.Empty );
          textDirection = tcPr.Element( textDirectionXName );
        }

        textDirection.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), value.ToString() );
      }
    }

    public int GridSpan
    {
      get
      {
        int gridSpanValue = 0;

        var tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        var gridSpan = tcPr?.Element( XName.Get( "gridSpan", Document.w.NamespaceName ) );
        if( gridSpan != null )
        {
          var gridSpanAttrValue = gridSpan.Attribute( XName.Get( "val", Document.w.NamespaceName ) );

          int value;
          if( gridSpanAttrValue != null && HelperFunctions.TryParseInt( gridSpanAttrValue.Value, out value ) )
            gridSpanValue = value;
        }
        return gridSpanValue;
      }
    }

    public int RowSpan
    {
      get
      {
        int rowSpanValue = 0;

        var tcPr = this.Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        var vMerge = tcPr?.Element( XName.Get( "vMerge", Document.w.NamespaceName ) );
        if( vMerge != null )
        {
          var vMergeAttrValue = vMerge.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
          // Starting a new vertical merge.
          if( ( vMergeAttrValue != null ) && ( vMergeAttrValue.Value == "restart" ) )
          {
            var rows = this._row._table.Rows;
            var rowIndex = rows.FindIndex( row => row.Xml == this._row.Xml );
            var cellIndex = this._row.Cells.FindIndex( cell => cell.Xml == this.Xml );
            if( ( rowIndex >= 0 ) && ( cellIndex >= 0 ) )
            {
              rowSpanValue = 1;

              for( var i = rowIndex + 1; i < rows.Count; ++i )
              {
                if( cellIndex >= rows[ i ].Cells.Count )
                  break;

                var cell = rows[ i ].Cells[ cellIndex ];
                var cell_tcPr = cell.Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
                var cell_vMerge = cell_tcPr?.Element( XName.Get( "vMerge", Document.w.NamespaceName ) );
                // vertical merge is done.
                if( cell_vMerge == null )
                  break;

                var cell_vMergeAttrValue = cell_vMerge.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
                // vertical merge is done, we are starting a new vMerge.
                if( ( cell_vMergeAttrValue != null ) && ( cell_vMergeAttrValue.Value == "restart" ) )
                  break;

                rowSpanValue++;
              }
            }
          }
        }
        return rowSpanValue;
      }
    }

    #endregion

    #region Constructors

    internal Cell( Row row, Document document, XElement xml )
        : base( document, xml )
    {
      _row = row;
      this.PackagePart = row.PackagePart;
    }



    #endregion

    #region Public Methods
















    // <summary>
    // Set the table cell border
    // </summary>
    // <example>
    // <code>
    // Create a new document.
    //using (var document = DocX.Create("Test.docx"))
    //{
    //    // Insert a table into this document.
    //    Table t = document.InsertTable(3, 3);
    //
    //    // Get the center cell.
    //    Cell center = t.Rows[1].Cells[1];
    //
    //    // Create a large blue border.
    //    Border b = new Border(BorderStyle.Tcbs_single, BorderSize.seven, 0, Color.Blue);
    //
    //    // Set the center cells Top, Bottom, Left and Right Borders to b.
    //    center.SetBorder(TableCellBorderType.Top, b);
    //    center.SetBorder(TableCellBorderType.Bottom, b);
    //    center.SetBorder(TableCellBorderType.Left, b);
    //    center.SetBorder(TableCellBorderType.Right, b);
    //
    //    // Save the document.
    //    document.Save();
    //}
    // </code>
    // </example>
    // <param name="borderType">Table Cell border to set</param>
    // <param name="border">Border object to set the table cell border</param>
    public void SetBorder( TableCellBorderType borderType, Border border )
    {
      /*
       * Get the tcPr (table cell properties) element for this Cell,
       * null will be return if no such element exists.
       */
      XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      if( tcPr == null )
      {
        Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
        tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      }

      /*
       * Get the tblBorders (table cell borders) element for this Cell,
       * null will be return if no such element exists.
       */
      XElement tcBorders = tcPr.Element( XName.Get( "tcBorders", Document.w.NamespaceName ) );
      if( tcBorders == null )
      {
        tcPr.SetElementValue( XName.Get( "tcBorders", Document.w.NamespaceName ), string.Empty );
        tcBorders = tcPr.Element( XName.Get( "tcBorders", Document.w.NamespaceName ) );
      }

      /*
       * Get the 'borderType' (table cell border) element for this Cell,
       * null will be return if no such element exists.
       */
      var tcbordertype = borderType.ToString();
      switch( borderType )
      {
        case TableCellBorderType.TopLeftToBottomRight:
          tcbordertype = "tl2br";
          break;
        case TableCellBorderType.TopRightToBottomLeft:
          tcbordertype = "tr2bl";
          break;
        default:
          // only lower the first char of string (because of insideH and insideV)
          tcbordertype = tcbordertype.Substring( 0, 1 ).ToLower() + tcbordertype.Substring( 1 );
          break;
      }

      XElement tcBorderType = tcBorders.Element( XName.Get( borderType.ToString(), Document.w.NamespaceName ) );
      if( tcBorderType == null )
      {
        tcBorders.SetElementValue( XName.Get( tcbordertype, Document.w.NamespaceName ), string.Empty );
        tcBorderType = tcBorders.Element( XName.Get( tcbordertype, Document.w.NamespaceName ) );
      }

      // get string value of border style
      string borderstyle = border.Tcbs.ToString().Substring( 5 );
      borderstyle = borderstyle.Substring( 0, 1 ).ToLower() + borderstyle.Substring( 1 );

      // The val attribute is used for the border style
      tcBorderType.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), borderstyle );

      int size;
      switch( border.Size )
      {
        case BorderSize.one:
          size = 2;
          break;
        case BorderSize.two:
          size = 4;
          break;
        case BorderSize.three:
          size = 6;
          break;
        case BorderSize.four:
          size = 8;
          break;
        case BorderSize.five:
          size = 12;
          break;
        case BorderSize.six:
          size = 18;
          break;
        case BorderSize.seven:
          size = 24;
          break;
        case BorderSize.eight:
          size = 36;
          break;
        case BorderSize.nine:
          size = 48;
          break;
        default:
          size = 2;
          break;
      }

      // The sz attribute is used for the border size
      tcBorderType.SetAttributeValue( XName.Get( "sz", Document.w.NamespaceName ), ( size ).ToString() );

      // The space attribute is used for the cell spacing (probably '0')
      tcBorderType.SetAttributeValue( XName.Get( "space", Document.w.NamespaceName ), ( border.Space ).ToString() );

      // The color attribute is used for the border color
      tcBorderType.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), border.Color.ToHex() );
    }

    public Border GetBorder( TableCellBorderType borderType )
    {
      // instance with default border values
      var b = new Border();

      /*
       * Get the tcPr (table cell properties) element for this Cell,
       * null will be return if no such element exists.
       */
      XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      if( tcPr == null )
      {
        // uses default border style
        return b;
      }

      /*
       * Get the tcBorders (table cell borders) element for this Cell,
       * null will be return if no such element exists.
       */
      XElement tcBorders = tcPr.Element( XName.Get( "tcBorders", Document.w.NamespaceName ) );
      if( tcBorders == null )
      {
        // uses default border style
        return b;
      }

      /*
       * Get the 'borderType' (cell border) element for this Cell,
       * null will be return if no such element exists.
       */
      var tcbordertype = borderType.ToString();
      switch( tcbordertype )
      {
        case "TopLeftToBottomRight":
          tcbordertype = "tl2br";
          break;
        case "TopRightToBottomLeft":
          tcbordertype = "tr2bl";
          break;
        default:
          // only lower the first char of string (because of insideH and insideV)
          tcbordertype = tcbordertype.Substring( 0, 1 ).ToLower() + tcbordertype.Substring( 1 );
          break;
      }

      XElement tcBorderType = tcBorders.Element( XName.Get( tcbordertype, Document.w.NamespaceName ) );
      if( tcBorderType == null )
      {
        // uses default border style
        return b;
      }

      // The val attribute is used for the border style
      XAttribute val = tcBorderType.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
      // If val is null, this cell contains no border information.
      if( val == null )
      {
        // uses default border style
      }
      else
      {
        try
        {
          string bordertype = "Tcbs_" + val.Value;
          b.Tcbs = ( BorderStyle )Enum.Parse( typeof( BorderStyle ), bordertype );
        }

        catch
        {
          val.Remove();
          // uses default border style
        }
      }

      // The sz attribute is used for the border size
      XAttribute sz = tcBorderType.Attribute( XName.Get( "sz", Document.w.NamespaceName ) );
      // If sz is null, this border contains no size information.
      if( sz == null )
      {
        // uses default border style
      }
      else
      {
        // If sz is not an int, something is wrong with this attributes value, so remove it
        int numerical_size;
        if( !HelperFunctions.TryParseInt( sz.Value, out numerical_size ) )
          sz.Remove();
        else
        {
          switch( numerical_size )
          {
            case 2:
              b.Size = BorderSize.one;
              break;
            case 4:
              b.Size = BorderSize.two;
              break;
            case 6:
              b.Size = BorderSize.three;
              break;
            case 8:
              b.Size = BorderSize.four;
              break;
            case 12:
              b.Size = BorderSize.five;
              break;
            case 18:
              b.Size = BorderSize.six;
              break;
            case 24:
              b.Size = BorderSize.seven;
              break;
            case 36:
              b.Size = BorderSize.eight;
              break;
            case 48:
              b.Size = BorderSize.nine;
              break;
            default:
              b.Size = BorderSize.one;
              break;
          }
        }
      }

      // The space attribute is used for the border spacing (probably '0')
      XAttribute space = tcBorderType.Attribute( XName.Get( "space", Document.w.NamespaceName ) );
      // If space is null, this border contains no space information.
      if( space == null )
      {
        // uses default border style
      }
      else
      {
        // If space is not a float, something is wrong with this attributes value, so remove it
        float borderspace;
        if( !HelperFunctions.TryParseFloat( space.Value, out borderspace ) )
        {
          space.Remove();
          // uses default border style
        }
        else
        {
          b.Space = borderspace;
        }
      }

      // The color attribute is used for the border color
      XAttribute color = tcBorderType.Attribute( XName.Get( "color", Document.w.NamespaceName ) );
      if( color == null )
      {
        // uses default border style
      }
      else
      {
        // If color is not a Color, something is wrong with this attributes value, so remove it
        try
        {
          b.Color = HelperFunctions.GetColorFromHtml( color.Value );
        }
        catch
        {
          color.Remove();
          // uses default border style
        }
      }
      return b;
    }

    public override Paragraph InsertParagraph( int index, Paragraph p )
    {
      this.RemoveDefaultParagraph();

      return base.InsertParagraph( index, p );
    }

    public override Paragraph InsertParagraph( int index, string text, bool trackChanges, Formatting formatting )
    {
      this.RemoveDefaultParagraph();

      return base.InsertParagraph( index, text, trackChanges, formatting );
    }

    public override Paragraph InsertParagraph( string text, bool trackChanges, Formatting formatting )
    {
      this.RemoveDefaultParagraph();

      return base.InsertParagraph( text, trackChanges, formatting );
    }

    public override Paragraph InsertParagraph( Paragraph p )
    {
      this.RemoveDefaultParagraph();

      return base.InsertParagraph( p );
    }

    public override Paragraph InsertEquation( string equation, Alignment align = Alignment.center )
    {
      this.RemoveDefaultParagraph();

      return base.InsertEquation( equation, align );
    }

    public override Table InsertTable( int rowCount, int columnCount )
    {
      this.RemoveDefaultParagraph();

      return base.InsertTable( rowCount, columnCount );
    }

    public override Table InsertTable( int index, int rowCount, int columnCount )
    {
      this.RemoveDefaultParagraph();

      return base.InsertTable( index, rowCount, columnCount );
    }

    public override Table InsertTable( Table t )
    {
      this.RemoveDefaultParagraph();

      return base.InsertTable( t );
    }

    public override Table InsertTable( int index, Table t )
    {
      this.RemoveDefaultParagraph();

      return base.InsertTable( index, t );
    }

    public override List InsertList( List list )
    {
      this.RemoveDefaultParagraph();

      return base.InsertList( list );
    }

    public override List InsertList( List list, double fontSize )
    {
      this.RemoveDefaultParagraph();

      return base.InsertList( list, fontSize );
    }

    public override List InsertList( List list, Font fontFamily, double fontSize )
    {
      this.RemoveDefaultParagraph();

      return base.InsertList( list, fontFamily, fontSize );
    }

    public override List InsertList( int index, List list )
    {
      this.RemoveDefaultParagraph();

      return base.InsertList( index, list );
    }


    #endregion

    #region Internal Methods

    internal void RemoveContentExcept( List<string> keepTexts )
    {
      var paragraphs = this.Paragraphs;
      var lastParagraphIdToKeep = paragraphs.Count;
      string lastParagraphTextToKeep = null;

      if( keepTexts != null )
      {
        for( int i = paragraphs.Count - 1; i >= 0; --i )
        {
          var paragraphText = paragraphs[ i ].Text;
          while( paragraphText.Contains( keepTexts.Last() ) )
          {
            paragraphText = paragraphText.Remove( paragraphText.LastIndexOf( keepTexts.Last() ) );
            keepTexts.RemoveAt( keepTexts.Count - 1 );

            // All this paragraph needs to be kept.
            if( string.IsNullOrEmpty( paragraphText ) || ( keepTexts.Count == 0 ) )
            {
              lastParagraphIdToKeep = i;
              break;
            }
          }

          if( keepTexts.Count == 0 )
          {
            // Some of this paragraph need to be kept.
            if( !string.IsNullOrEmpty( paragraphText ) )
            {
              lastParagraphTextToKeep = paragraphText;
            }

            break;
          }
        }
      }

      // Remove unkept text from last paragraph to keep.
      if( lastParagraphTextToKeep != null )
      {
        this.Paragraphs[ lastParagraphIdToKeep ].RemoveText( 0, lastParagraphTextToKeep.Length );
      }

      // Remove unkept first paragraphs.
      for( int i = lastParagraphIdToKeep - 1; i >= 0; i-- )
      {
        this.Paragraphs[ i ].Remove( false );
      }
    }

    #endregion

    #region Private Methods

    private void RemoveDefaultParagraph()
    {
      if( this.IsDefaultParagraph() )
      {
        this.Paragraphs[ 0 ].Remove( false );
      }
    }

    private bool IsDefaultParagraph()
    {
      if( this.Paragraphs.Count < 1 )
        return false;

      var firstParagraph = this.Paragraphs[ 0 ];
      return this.Paragraphs.Count == 1
             && string.IsNullOrEmpty( firstParagraph.Text )
             && !firstParagraph.IsListItem
             && firstParagraph.Pictures.Count == 0
             && firstParagraph.Hyperlinks.Count == 0;

    }






































    private void UpdateShadingPatternXml()
    {
      if( _shadingPattern == null )
      {
        throw new Exception( "Shading pattern value is invalid." );
      }

      /*
       * Get the tcPr (table cell properties) element for this Cell,
       * null will be return if no such element exists.
       */
      XElement tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      if( tcPr == null )
      {
        Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
        tcPr = Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      }

      /*
       * Get the shd (table shade) element for this Cell,
       * null will be return if no such element exists.
       */
      XElement shd = tcPr.Element( XName.Get( "shd", Document.w.NamespaceName ) );
      if( shd == null )
      {
        tcPr.SetElementValue( XName.Get( "shd", Document.w.NamespaceName ), string.Empty );
        shd = tcPr.Element( XName.Get( "shd", Document.w.NamespaceName ) );
      }

      // The fill attribute needs to be set to the hex for this Color.
      shd.SetAttributeValue( XName.Get( "fill", Document.w.NamespaceName ), _shadingPattern.Fill.ToHex() );

      // Set the value to val attribute.
      shd.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), HelperFunctions.GetValueFromTablePatternStyle( _shadingPattern.Style ) );

      // The color attribute needs to be set to the hex for this Color.
      shd.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), _shadingPattern.Style == PatternStyle.Clear ? "auto" : _shadingPattern.StyleColor.ToHex() );
    }

    #endregion

    #region Event Handlers

    private void ShadingPattern_PropertyChanged( object sender, PropertyChangedEventArgs e )
    {
      this.UpdateShadingPatternXml();
    }

    #endregion
  }
}
