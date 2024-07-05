/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2024 Xceed Software Inc.
 
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
using System.IO.Packaging;
using System.IO;
using System.Drawing;
using System.Globalization;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;

namespace Xceed.Document.NET
{
  public class Table : InsertBeforeOrAfter
  {
    #region Private Members

    private Alignment _alignment;
    private AutoFit _autofit;
    private TableDesign _design;
    private TableLook _tableLook;
    private double _indentFromLeft;

    private string _customTableDesignName;
    private int _cachedColumnCount = -1;

    #endregion

    #region Public Properties

    public virtual List<Paragraph> Paragraphs
    {
      get
      {
        var paragraphs = new List<Paragraph>();

        foreach( Row r in Rows )
          paragraphs.AddRange( r.Paragraphs );

        return paragraphs;
      }
    }

    public List<Picture> Pictures
    {
      get
      {
        var pictures = new List<Picture>();

        foreach( Row r in Rows )
        {
          pictures.AddRange( r.Pictures );
        }

        return pictures;
      }
    }







    public List<Hyperlink> Hyperlinks
    {
      get
      {
        var hyperlinks = new List<Hyperlink>();

        foreach( Row r in Rows )
        {
          hyperlinks.AddRange( r.Hyperlinks );
        }

        return hyperlinks;
      }
    }

    public Int32 RowCount
    {
      get
      {
        return this.Xml.Elements( XName.Get( "tr", Document.w.NamespaceName ) ).Count();
      }
    }

    public Int32 ColumnCount
    {
      get
      {
        if( this.RowCount == 0 )
          return 0;
        if( _cachedColumnCount == -1 )
        {
          foreach( var r in this.Rows )
          {
            _cachedColumnCount = Math.Max( _cachedColumnCount, r.ColumnCount );
          }
        }
        return _cachedColumnCount;
      }
    }

    public List<Row> Rows
    {
      get
      {
        var rows =
        (
            from r in Xml.Elements( XName.Get( "tr", Document.w.NamespaceName ) )
            select new Row( this, Document, r )
        ).ToList();

        return rows;
      }
    }

    public Alignment Alignment
    {
      get
      {
        return _alignment;
      }
      set
      {
        string alignmentString = string.Empty;
        switch( value )
        {
          case Alignment.left:
            {
              alignmentString = "left";
              break;
            }

          case Alignment.both:
            {
              alignmentString = "both";
              break;
            }

          case Alignment.right:
            {
              alignmentString = "right";
              break;
            }

          case Alignment.center:
            {
              alignmentString = "center";
              break;
            }
        }

        XElement tblPr = Xml.Descendants( XName.Get( "tblPr", Document.w.NamespaceName ) ).First();
        XElement jc = tblPr.Descendants( XName.Get( "jc", Document.w.NamespaceName ) ).FirstOrDefault();

        jc?.Remove();

        jc = new XElement( XName.Get( "jc", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), alignmentString ) );
        tblPr.Add( jc );
        _alignment = value;
      }
    }

    public double IndentFromLeft
    {
      get
      {
        return _indentFromLeft / 20d;
      }
      set
      {
        _indentFromLeft = value * 20d;

        var tblPr = this.Xml.Descendants( XName.Get( "tblPr", Document.w.NamespaceName ) ).First();
        var tblInd = tblPr.Element( XName.Get( "tblInd", Document.w.NamespaceName ) );
        if( tblInd == null )
        {
          tblPr.Add( new XElement( XName.Get( "tblInd", Document.w.NamespaceName ) ) );
          tblInd = tblPr.Element( XName.Get( "tblInd", Document.w.NamespaceName ) );
        }
        tblInd.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), _indentFromLeft );
        tblInd.Add( new XAttribute( XName.Get( "type", Document.w.NamespaceName ), "dxa" ) );
      }
    }

    public AutoFit AutoFit
    {
      get
      {
        return _autofit;
      }

      set
      {
        string tableAttributeValue = string.Empty;
        string columnAttributeValue = string.Empty;
        switch( value )
        {
          case AutoFit.ColumnWidth:
            {
              tableAttributeValue = "auto";
              columnAttributeValue = "dxa";

              // Disable "Automatically resize to fit contents" option
              var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
              if( tblPr != null )
              {
                var layout = tblPr.Element( XName.Get( "tblLayout", Document.w.NamespaceName ) );
                if( layout == null )
                {
                  tblPr.Add( new XElement( XName.Get( "tblLayout", Document.w.NamespaceName ) ) );
                  layout = tblPr.Element( XName.Get( "tblLayout", Document.w.NamespaceName ) );
                }

                var type = layout.Attribute( XName.Get( "type", Document.w.NamespaceName ) );
                if( type == null )
                {
                  layout.Add( new XAttribute( XName.Get( "type", Document.w.NamespaceName ), String.Empty ) );
                  type = layout.Attribute( XName.Get( "type", Document.w.NamespaceName ) );
                }

                type.Value = "fixed";
              }

              break;
            }

          case AutoFit.Contents:
            {
              tableAttributeValue = columnAttributeValue = "auto";

              // Set table tblW to 0.
              this.UpdateTableWidth( AutoFit.Contents, "0" );

              // Set table width type to auto.
              var tblPr = this.Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );

              if( tblPr != null )
              {
                var tblLayout = tblPr.Element( XName.Get( "tblLayout", Document.w.NamespaceName ) );

                if( tblLayout == null )
                {
                  var tmp = tblPr.Element( XName.Get( "tblW", Document.w.NamespaceName ) );

                  tmp.AddAfterSelf( new XElement( XName.Get( "tblLayout", Document.w.NamespaceName ) ) );
                  tmp = tblPr.Element( XName.Get( "tblLayout", Document.w.NamespaceName ) );
                  tmp.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), "autofit" );
                  tmp = tblPr.Element( XName.Get( "tblW", Document.w.NamespaceName ) );
                  tmp.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), tableAttributeValue );
                }
                else
                {
                  var types = from d in Xml.Descendants()
                              let type = d.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
                              where ( d.Name.LocalName == "tblLayout" ) && type != null
                              select type;

                  foreach( XAttribute type in types )
                  {
                    type.Value = "autofit";
                  }

                  var tmp = tblPr.Element( XName.Get( "tblW", Document.w.NamespaceName ) );
                  tmp.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), "auto" );
                }
              }
              // Set table cells tcW to 0.
              var tcW = from d in this.Xml.Descendants()
                        let type = d.Attribute( XName.Get( "w", Document.w.NamespaceName ) )
                        where ( d.Name.LocalName == "tcW" ) && ( type != null )
                        select type;

              foreach( var w in tcW )
              {
                w.Value = "0";
              }

              break;
            }

          case AutoFit.Window:
            {
              tableAttributeValue = columnAttributeValue = "pct";

              // Set table width to 5000 and width type to percentage.
              this.UpdateTableWidth( AutoFit.Window, "5000" );

              // Remove table indentation and layout properties
              var tblPr = this.Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );

              if( tblPr != null )
              {
                var tblInd = tblPr.Element( XName.Get( "tblInd", Document.w.NamespaceName ) );
                if( tblInd != null )
                {
                  tblInd.Remove();
                }

                var tblLayout = tblPr.Element( XName.Get( "tblLayout", Document.w.NamespaceName ) );
                if( tblLayout != null )
                {
                  tblLayout.Remove();
                }
              }

              break;
            }

          case AutoFit.Fixed:
            {
              tableAttributeValue = columnAttributeValue = "dxa";
              var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
              var tblLayout = tblPr.Element( XName.Get( "tblLayout", Document.w.NamespaceName ) );

              if( tblLayout == null )
              {
                var tmp = tblPr.Element( XName.Get( "tblInd", Document.w.NamespaceName ) ) ?? tblPr.Element( XName.Get( "tblW", Document.w.NamespaceName ) );

                tmp.AddAfterSelf( new XElement( XName.Get( "tblLayout", Document.w.NamespaceName ) ) );
                tmp = tblPr.Element( XName.Get( "tblLayout", Document.w.NamespaceName ) );
                tmp.SetAttributeValue( XName.Get( "type", Document.w.NamespaceName ), "fixed" );
                tmp = tblPr.Element( XName.Get( "tblW", Document.w.NamespaceName ) );

                Double totalWidth = 0;
                if( this.ColumnWidths != null )
                {
                  foreach( Double columnWidth in ColumnWidths )
                  {
                    totalWidth += columnWidth;
                  }
                }

                tmp.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), ( totalWidth * 20d ).ToString( CultureInfo.InvariantCulture ) );
                break;
              }
              else
              {
                var types = from d in Xml.Descendants()
                            let type = d.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
                            where ( d.Name.LocalName == "tblLayout" ) && type != null
                            select type;

                foreach( XAttribute type in types )
                {
                  type.Value = "fixed";
                }

                var tmp = tblPr.Element( XName.Get( "tblW", Document.w.NamespaceName ) );

                Double totalWidth = 0;
                if( this.ColumnWidths != null )
                {
                  foreach( Double columnWidth in this.ColumnWidths )
                  {
                    totalWidth += columnWidth;
                  }
                }

                tmp.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), ( totalWidth * 20d ).ToString( CultureInfo.InvariantCulture ) );
                break;
              }
            }
        }

        if( value != AutoFit.Window )
        {
          // Set table attributes
          var query = from d in Xml.Descendants()
                      let type = d.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
                      where ( d.Name.LocalName == "tblW" ) && type != null
                      select type;

          foreach( XAttribute type in query )
          {
            type.Value = tableAttributeValue;
          }

          // Set column attributes
          query = from d in Xml.Descendants()
                  let type = d.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
                  where ( d.Name.LocalName == "tcW" ) && type != null
                  select type;

          foreach( XAttribute type in query )
          {
            type.Value = columnAttributeValue;
          }
        }

        _autofit = value;
      }
    }

    public TableDesign Design
    {
      get
      {
        return _design;
      }
      set
      {
        if( _design != value )
        {
          XElement tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
          XElement style = tblPr.Element( XName.Get( "tblStyle", Document.w.NamespaceName ) );
          if( style == null )
          {
            tblPr.Add( new XElement( XName.Get( "tblStyle", Document.w.NamespaceName ) ) );
            style = tblPr.Element( XName.Get( "tblStyle", Document.w.NamespaceName ) );
          }

          XAttribute val = style.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
          if( val == null )
          {
            style.Add( new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "" ) );
            val = style.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
          }

          _design = value;

          if( _design == TableDesign.None )
          {
            if( style != null )
              style.Remove();
          }

          if( _design == TableDesign.Custom )
          {
            if( string.IsNullOrEmpty( _customTableDesignName ) )
            {
              _design = TableDesign.None;
              if( style != null )
                style.Remove();

            }
            else
            {
              val.Value = _customTableDesignName;
            }
          }
          else
          {
            switch( _design )
            {
              case TableDesign.TableNormal:
                val.Value = "TableNormal";
                break;
              case TableDesign.TableGrid:
                val.Value = "TableGrid";
                break;
              case TableDesign.LightShading:
                val.Value = "LightShading";
                break;
              case TableDesign.LightShadingAccent1:
                val.Value = "LightShading-Accent1";
                break;
              case TableDesign.LightShadingAccent2:
                val.Value = "LightShading-Accent2";
                break;
              case TableDesign.LightShadingAccent3:
                val.Value = "LightShading-Accent3";
                break;
              case TableDesign.LightShadingAccent4:
                val.Value = "LightShading-Accent4";
                break;
              case TableDesign.LightShadingAccent5:
                val.Value = "LightShading-Accent5";
                break;
              case TableDesign.LightShadingAccent6:
                val.Value = "LightShading-Accent6";
                break;
              case TableDesign.LightList:
                val.Value = "LightList";
                break;
              case TableDesign.LightListAccent1:
                val.Value = "LightList-Accent1";
                break;
              case TableDesign.LightListAccent2:
                val.Value = "LightList-Accent2";
                break;
              case TableDesign.LightListAccent3:
                val.Value = "LightList-Accent3";
                break;
              case TableDesign.LightListAccent4:
                val.Value = "LightList-Accent4";
                break;
              case TableDesign.LightListAccent5:
                val.Value = "LightList-Accent5";
                break;
              case TableDesign.LightListAccent6:
                val.Value = "LightList-Accent6";
                break;
              case TableDesign.LightGrid:
                val.Value = "LightGrid";
                break;
              case TableDesign.LightGridAccent1:
                val.Value = "LightGrid-Accent1";
                break;
              case TableDesign.LightGridAccent2:
                val.Value = "LightGrid-Accent2";
                break;
              case TableDesign.LightGridAccent3:
                val.Value = "LightGrid-Accent3";
                break;
              case TableDesign.LightGridAccent4:
                val.Value = "LightGrid-Accent4";
                break;
              case TableDesign.LightGridAccent5:
                val.Value = "LightGrid-Accent5";
                break;
              case TableDesign.LightGridAccent6:
                val.Value = "LightGrid-Accent6";
                break;
              case TableDesign.MediumShading1:
                val.Value = "MediumShading1";
                break;
              case TableDesign.MediumShading1Accent1:
                val.Value = "MediumShading1-Accent1";
                break;
              case TableDesign.MediumShading1Accent2:
                val.Value = "MediumShading1-Accent2";
                break;
              case TableDesign.MediumShading1Accent3:
                val.Value = "MediumShading1-Accent3";
                break;
              case TableDesign.MediumShading1Accent4:
                val.Value = "MediumShading1-Accent4";
                break;
              case TableDesign.MediumShading1Accent5:
                val.Value = "MediumShading1-Accent5";
                break;
              case TableDesign.MediumShading1Accent6:
                val.Value = "MediumShading1-Accent6";
                break;
              case TableDesign.MediumShading2:
                val.Value = "MediumShading2";
                break;
              case TableDesign.MediumShading2Accent1:
                val.Value = "MediumShading2-Accent1";
                break;
              case TableDesign.MediumShading2Accent2:
                val.Value = "MediumShading2-Accent2";
                break;
              case TableDesign.MediumShading2Accent3:
                val.Value = "MediumShading2-Accent3";
                break;
              case TableDesign.MediumShading2Accent4:
                val.Value = "MediumShading2-Accent4";
                break;
              case TableDesign.MediumShading2Accent5:
                val.Value = "MediumShading2-Accent5";
                break;
              case TableDesign.MediumShading2Accent6:
                val.Value = "MediumShading2-Accent6";
                break;
              case TableDesign.MediumList1:
                val.Value = "MediumList1";
                break;
              case TableDesign.MediumList1Accent1:
                val.Value = "MediumList1-Accent1";
                break;
              case TableDesign.MediumList1Accent2:
                val.Value = "MediumList1-Accent2";
                break;
              case TableDesign.MediumList1Accent3:
                val.Value = "MediumList1-Accent3";
                break;
              case TableDesign.MediumList1Accent4:
                val.Value = "MediumList1-Accent4";
                break;
              case TableDesign.MediumList1Accent5:
                val.Value = "MediumList1-Accent5";
                break;
              case TableDesign.MediumList1Accent6:
                val.Value = "MediumList1-Accent6";
                break;
              case TableDesign.MediumList2:
                val.Value = "MediumList2";
                break;
              case TableDesign.MediumList2Accent1:
                val.Value = "MediumList2-Accent1";
                break;
              case TableDesign.MediumList2Accent2:
                val.Value = "MediumList2-Accent2";
                break;
              case TableDesign.MediumList2Accent3:
                val.Value = "MediumList2-Accent3";
                break;
              case TableDesign.MediumList2Accent4:
                val.Value = "MediumList2-Accent4";
                break;
              case TableDesign.MediumList2Accent5:
                val.Value = "MediumList2-Accent5";
                break;
              case TableDesign.MediumList2Accent6:
                val.Value = "MediumList2-Accent6";
                break;
              case TableDesign.MediumGrid1:
                val.Value = "MediumGrid1";
                break;
              case TableDesign.MediumGrid1Accent1:
                val.Value = "MediumGrid1-Accent1";
                break;
              case TableDesign.MediumGrid1Accent2:
                val.Value = "MediumGrid1-Accent2";
                break;
              case TableDesign.MediumGrid1Accent3:
                val.Value = "MediumGrid1-Accent3";
                break;
              case TableDesign.MediumGrid1Accent4:
                val.Value = "MediumGrid1-Accent4";
                break;
              case TableDesign.MediumGrid1Accent5:
                val.Value = "MediumGrid1-Accent5";
                break;
              case TableDesign.MediumGrid1Accent6:
                val.Value = "MediumGrid1-Accent6";
                break;
              case TableDesign.MediumGrid2:
                val.Value = "MediumGrid2";
                break;
              case TableDesign.MediumGrid2Accent1:
                val.Value = "MediumGrid2-Accent1";
                break;
              case TableDesign.MediumGrid2Accent2:
                val.Value = "MediumGrid2-Accent2";
                break;
              case TableDesign.MediumGrid2Accent3:
                val.Value = "MediumGrid2-Accent3";
                break;
              case TableDesign.MediumGrid2Accent4:
                val.Value = "MediumGrid2-Accent4";
                break;
              case TableDesign.MediumGrid2Accent5:
                val.Value = "MediumGrid2-Accent5";
                break;
              case TableDesign.MediumGrid2Accent6:
                val.Value = "MediumGrid2-Accent6";
                break;
              case TableDesign.MediumGrid3:
                val.Value = "MediumGrid3";
                break;
              case TableDesign.MediumGrid3Accent1:
                val.Value = "MediumGrid3-Accent1";
                break;
              case TableDesign.MediumGrid3Accent2:
                val.Value = "MediumGrid3-Accent2";
                break;
              case TableDesign.MediumGrid3Accent3:
                val.Value = "MediumGrid3-Accent3";
                break;
              case TableDesign.MediumGrid3Accent4:
                val.Value = "MediumGrid3-Accent4";
                break;
              case TableDesign.MediumGrid3Accent5:
                val.Value = "MediumGrid3-Accent5";
                break;
              case TableDesign.MediumGrid3Accent6:
                val.Value = "MediumGrid3-Accent6";
                break;

              case TableDesign.DarkList:
                val.Value = "DarkList";
                break;
              case TableDesign.DarkListAccent1:
                val.Value = "DarkList-Accent1";
                break;
              case TableDesign.DarkListAccent2:
                val.Value = "DarkList-Accent2";
                break;
              case TableDesign.DarkListAccent3:
                val.Value = "DarkList-Accent3";
                break;
              case TableDesign.DarkListAccent4:
                val.Value = "DarkList-Accent4";
                break;
              case TableDesign.DarkListAccent5:
                val.Value = "DarkList-Accent5";
                break;
              case TableDesign.DarkListAccent6:
                val.Value = "DarkList-Accent6";
                break;

              case TableDesign.ColorfulShading:
                val.Value = "ColorfulShading";
                break;
              case TableDesign.ColorfulShadingAccent1:
                val.Value = "ColorfulShading-Accent1";
                break;
              case TableDesign.ColorfulShadingAccent2:
                val.Value = "ColorfulShading-Accent2";
                break;
              case TableDesign.ColorfulShadingAccent3:
                val.Value = "ColorfulShading-Accent3";
                break;
              case TableDesign.ColorfulShadingAccent4:
                val.Value = "ColorfulShading-Accent4";
                break;
              case TableDesign.ColorfulShadingAccent5:
                val.Value = "ColorfulShading-Accent5";
                break;
              case TableDesign.ColorfulShadingAccent6:
                val.Value = "ColorfulShading-Accent6";
                break;

              case TableDesign.ColorfulList:
                val.Value = "ColorfulList";
                break;
              case TableDesign.ColorfulListAccent1:
                val.Value = "ColorfulList-Accent1";
                break;
              case TableDesign.ColorfulListAccent2:
                val.Value = "ColorfulList-Accent2";
                break;
              case TableDesign.ColorfulListAccent3:
                val.Value = "ColorfulList-Accent3";
                break;
              case TableDesign.ColorfulListAccent4:
                val.Value = "ColorfulList-Accent4";
                break;
              case TableDesign.ColorfulListAccent5:
                val.Value = "ColorfulList-Accent5";
                break;
              case TableDesign.ColorfulListAccent6:
                val.Value = "ColorfulList-Accent6";
                break;

              case TableDesign.ColorfulGrid:
                val.Value = "ColorfulGrid";
                break;
              case TableDesign.ColorfulGridAccent1:
                val.Value = "ColorfulGrid-Accent1";
                break;
              case TableDesign.ColorfulGridAccent2:
                val.Value = "ColorfulGrid-Accent2";
                break;
              case TableDesign.ColorfulGridAccent3:
                val.Value = "ColorfulGrid-Accent3";
                break;
              case TableDesign.ColorfulGridAccent4:
                val.Value = "ColorfulGrid-Accent4";
                break;
              case TableDesign.ColorfulGridAccent5:
                val.Value = "ColorfulGrid-Accent5";
                break;
              case TableDesign.ColorfulGridAccent6:
                val.Value = "ColorfulGrid-Accent6";
                break;

              default:
                break;
            }
          }

          if( Document._styles == null )
          {
            PackagePart word_styles = Document._package.GetPart( new Uri( "/word/styles.xml", UriKind.Relative ) );
            using( TextReader tr = new StreamReader( word_styles.GetStream() ) )
              Document._styles = XDocument.Load( tr );
          }

          if( !string.IsNullOrEmpty( val.Value ) )
          {
            var tableStyle =
            (
                from e in Document._styles.Descendants()
                let styleId = e.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) )
                where ( styleId != null && styleId.Value == val.Value )
                select e
            ).FirstOrDefault();

            if( tableStyle == null )
            {
              XDocument external_style_doc = HelperFunctions.DecompressXMLResource( HelperFunctions.GetResources( ResourceType.Styles ) );

              var styleElement =
              (
                  from e in external_style_doc.Descendants()
                  let styleId = e.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) )
                  where ( styleId != null && styleId.Value == val.Value )
                  select e
              ).FirstOrDefault();

              if( styleElement != null )
                Document._styles.Element( XName.Get( "styles", Document.w.NamespaceName ) ).Add( styleElement );
            }
          }
        }
      }
    }

    public int Index
    {
      get
      {
        int index = 0;
        IEnumerable<XElement> previous = Xml.ElementsBeforeSelf();

        foreach( XElement e in previous )
          index += Paragraph.GetElementTextLength( e );

        return index;
      }
    }

    public string CustomTableDesignName
    {
      get
      {
        return _customTableDesignName;
      }
      set
      {
        _customTableDesignName = value;
        this.Design = TableDesign.Custom;
      }

    }

    public string TableCaption
    {
      get
      {
        var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
        var caption = tblPr?.Element( XName.Get( "tblCaption", Document.w.NamespaceName ) );

        if( caption != null )
          return caption.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );

        return null;
      }
      set
      {
        var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
        if( tblPr != null )
        {
          var caption = tblPr.Descendants( XName.Get( "tblCaption", Document.w.NamespaceName ) ).FirstOrDefault();
          if( caption != null )
          {
            caption.Remove();
          }
          caption = new XElement( XName.Get( "tblCaption", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), value ) );
          tblPr.Add( caption );
        }
      }
    }

    public string TableDescription
    {
      get
      {
        var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
        var description = tblPr?.Element( XName.Get( "tblDescription", Document.w.NamespaceName ) );

        if( description != null )
          return description.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );

        return null;
      }
      set
      {
        var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
        if( tblPr != null )
        {
          var description = tblPr.Descendants( XName.Get( "tblDescription", Document.w.NamespaceName ) ).FirstOrDefault();
          description?.Remove();
          description = new XElement( XName.Get( "tblDescription", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), value ) );
          tblPr.Add( description );
        }
      }
    }


    public TableLook TableLook
    {
      get
      {
        return _tableLook;
      }
      set
      {
        if( _tableLook != null )
        {
          _tableLook.PropertyChanged -= this.TableLook_PropertyChanged;
        }
        _tableLook = value;
        _tableLook.PropertyChanged += this.TableLook_PropertyChanged;

        this.UpdateTableLookXml();
      }
    }

    public List<Double> ColumnWidths
    {
      get
      {
        var columnWidths = new List<Double>();

        // get the table grid property
        XElement grid = Xml.Element( XName.Get( "tblGrid", Document.w.NamespaceName ) );

        // get the columns properties
        var columns = grid?.Elements( XName.Get( "gridCol", Document.w.NamespaceName ) );
        if( columns == null )
          return null;

        foreach( var column in columns )
        {
          string value = column.GetAttribute( XName.Get( "w", Document.w.NamespaceName ) );
          columnWidths.Add( Convert.ToDouble( value, new CultureInfo( "en-US" ) ) / 20d );
        }

        return columnWidths;
      }
    }

    public ShadingPattern ShadingPattern
    {
      get
      {
        if( Rows.Count < 0 )
          throw new IndexOutOfRangeException();

        if( Rows[ 0 ].ColumnCount < 0 )
          throw new IndexOutOfRangeException();

        ShadingPattern shadingPattern = Rows[ 0 ].Cells[ 0 ].ShadingPattern;

        foreach( Row r in Rows )
        {
          var cells = r.Cells;

          foreach( Cell c in cells )
          {
            if( !shadingPattern.Equals( c.ShadingPattern ) )
            {
              return new ShadingPattern();
            }
          }
        }

        return shadingPattern;
      }

      set
      {
        foreach( Row r in Rows )
        {
          var cells = r.Cells;

          foreach( Cell c in cells )
          {
            c.ShadingPattern = value;
          }
        }
      }
    }

    #endregion

    #region Constructors

    internal Table( Document document, XElement xml, PackagePart packagePart )
        : base( document, xml )
    {
      _autofit = Table.GetAutoFitFromXml( xml );
      this.Xml = xml;
      this.PackagePart = packagePart;

      var properties = xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );

      var tblGrid = xml.Element( XName.Get( "tblGrid", Document.w.NamespaceName ) );
      if( tblGrid == null )
      {
        this.SetColumnWidth( 0, -1, false );
      }

      var alignment = properties?.Element( XName.Get( "jc", Document.w.NamespaceName ) );
      if( alignment != null )
      {
        var val = alignment.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
        if( val != null )
        {
          _alignment = ( Alignment )Enum.Parse( typeof( Alignment ), val.Value );
        }
      }

      var tblInd = properties?.Element( XName.Get( "tblInd", Document.w.NamespaceName ) );
      if( tblInd != null )
      {
        var val = tblInd.Attribute( XName.Get( "w", Document.w.NamespaceName ) );
        if( val != null )
        {
          _indentFromLeft = double.Parse( val.Value, CultureInfo.InvariantCulture );
        }
      }

      var style = properties?.Element( XName.Get( "tblStyle", Document.w.NamespaceName ) );
      if( style != null )
      {
        var val = style.Attribute( XName.Get( "val", Document.w.NamespaceName ) );

        if( val != null )
        {
          String cleanValue = val.Value.Replace( "-", string.Empty );

          if( Enum.IsDefined( typeof( TableDesign ), cleanValue ) )
          {
            this.Design = ( TableDesign )Enum.Parse( typeof( TableDesign ), cleanValue );
          }

          else
          {
            this.Design = TableDesign.Custom;
            this.CustomTableDesignName = val.Value;
          }
        }
        else
        {
          this.Design = TableDesign.None;
        }
      }
      else
      {
        this.Design = TableDesign.None;
      }

      var tableLook = properties?.Element( XName.Get( "tblLook", Document.w.NamespaceName ) );
      if( tableLook != null )
      {
        // Using "val" is the old way of setting the tableLook.
        var val = tableLook.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );
        if( !string.IsNullOrEmpty( val ) )
        {
          var tableLookValue = Int32.Parse( val, NumberStyles.HexNumber );
          this.TableLook = new TableLook( ( tableLookValue & Int32.Parse( "0020", NumberStyles.HexNumber ) ) != 0,
                                          ( tableLookValue & Int32.Parse( "0040", NumberStyles.HexNumber ) ) != 0,
                                          ( tableLookValue & Int32.Parse( "0080", NumberStyles.HexNumber ) ) != 0,
                                          ( tableLookValue & Int32.Parse( "0100", NumberStyles.HexNumber ) ) != 0,
                                          ( tableLookValue & Int32.Parse( "0200", NumberStyles.HexNumber ) ) != 0,
                                          ( tableLookValue & Int32.Parse( "0400", NumberStyles.HexNumber ) ) != 0 );
        }
        else
        {
          this.TableLook = new TableLook( tableLook.GetAttribute( XName.Get( "firstRow", Document.w.NamespaceName ) ) == "1",
                                          tableLook.GetAttribute( XName.Get( "lastRow", Document.w.NamespaceName ) ) == "1",
                                          tableLook.GetAttribute( XName.Get( "firstColumn", Document.w.NamespaceName ) ) == "1",
                                          tableLook.GetAttribute( XName.Get( "lastColumn", Document.w.NamespaceName ) ) == "1",
                                          tableLook.GetAttribute( XName.Get( "noHBand", Document.w.NamespaceName ) ) == "1",
                                          tableLook.GetAttribute( XName.Get( "noVBand", Document.w.NamespaceName ) ) == "1" );
        }
      }

    }

    #endregion

    #region Public Methods

    public void MergeCellsInColumn( int columnIndex, int startRow, int endRow )
    {
      // Check for valid start and end indexes.
      if( columnIndex < 0 || columnIndex >= ColumnCount )
        throw new IndexOutOfRangeException();

      if( startRow < 0 || endRow <= startRow || endRow >= Rows.Count )
        throw new IndexOutOfRangeException();
      // Foreach each Cell between startIndex and endIndex inclusive.
      var validRows = this.Rows.GetRange( startRow, endRow - startRow + 1 );
      foreach( var row in validRows )
      {
        var c = row.Cells[ columnIndex ];
        var tcPr = c.Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        if( tcPr == null )
        {
          c.Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
          tcPr = c.Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
        }

        var vMerge = tcPr.Element( XName.Get( "vMerge", Document.w.NamespaceName ) );
        if( vMerge == null )
        {
          tcPr.SetElementValue( XName.Get( "vMerge", Document.w.NamespaceName ), string.Empty );
          vMerge = tcPr.Element( XName.Get( "vMerge", Document.w.NamespaceName ) );
        }
      }

      /* 
       * Get the tcPr (table cell properties) element for the first cell in this merge,
      * null will be returned if no such element exists.
       */
      var startRowCellsCount = this.Rows[ startRow ].Cells.Count;
      var start_tcPr = ( columnIndex > startRowCellsCount )
                       ? this.Rows[ startRow ].Cells[ startRowCellsCount - 1 ].Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) )
                       : this.Rows[ startRow ].Cells[ columnIndex ].Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      if( start_tcPr == null )
      {
        this.Rows[ startRow ].Cells[ columnIndex ].Xml.SetElementValue( XName.Get( "tcPr", Document.w.NamespaceName ), string.Empty );
        start_tcPr = this.Rows[ startRow ].Cells[ columnIndex ].Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
      }

      /* 
        * Get the gridSpan element of this row,
        * null will be returned if no such element exists.
        */
      var start_vMerge = start_tcPr.Element( XName.Get( "vMerge", Document.w.NamespaceName ) );
      if( start_vMerge == null )
      {
        start_tcPr.SetElementValue( XName.Get( "vMerge", Document.w.NamespaceName ), string.Empty );
        start_vMerge = start_tcPr.Element( XName.Get( "vMerge", Document.w.NamespaceName ) );
      }

      start_vMerge.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), "restart" );
    }

    public void SetDirection( Direction direction )
    {
      var tblPr = GetOrCreate_tblPr();
      tblPr.Add( new XElement( Document.w + "bidiVisual" ) );

      foreach( Row r in Rows )
      {
        r.SetDirection( direction );
      }
    }

    public void Remove()
    {
      this.RemoveInternal();

      // Update Sections tables.
      this.Document.UpdateCacheSections();
    }

    public Row InsertRow()
    {
      return this.InsertRow( this.RowCount );
    }

    public Row InsertRow( Row row, bool keepFormatting = false )
    {
      return this.InsertRow( row, this.RowCount, keepFormatting );
    }

    public void InsertColumn()
    {
      this.InsertColumn( this.ColumnCount - 1, true );
    }

    public void RemoveRow()
    {
      this.RemoveRow( RowCount - 1 );
    }

    public void RemoveRow( int index )
    {
      if( index < 0 || index > RowCount - 1 )
        throw new IndexOutOfRangeException();

      this.Document.ClearParagraphsCache();

      this.Rows[ index ].Xml.Remove();
      if( this.Rows.Count == 0 )
      {
        this.Remove();
      }
    }

    public void RemoveColumn()
    {
      this.RemoveColumn( this.ColumnCount - 1 );
    }

    public void RemoveColumn( int index )
    {
      if( ( index < 0 ) || ( index > this.ColumnCount - 1 ) )
        throw new IndexOutOfRangeException();

      foreach( Row r in Rows )
      {
        if( r.Cells.Count < this.ColumnCount )
        {
          int gridAfterValue = r.GridAfter;
          int posIndex = 0;
          int currentPos = 0;

          for( int i = 0; i < r.Cells.Count; ++i )
          {
            var rowCell = r.Cells[ i ];

            int gridSpanValue = ( rowCell.GridSpan != 0 ) ? rowCell.GridSpan - 1 : 0;

            // checks to see if index is between the lowest and highest cell value.
            if( ( ( index - gridAfterValue ) >= currentPos )
              && ( ( index - gridAfterValue ) <= ( currentPos + gridSpanValue ) ) )
            {
              r.Cells[ posIndex ].Xml.Remove();
              break;
            }

            ++posIndex;
            currentPos += ( gridSpanValue + 1 );
          }
        }
        else
        {
          r.Cells[ index ].Xml.Remove();
        }
      }
      _cachedColumnCount = -1;

      this.Document.ClearParagraphsCache();
    }

    public Row InsertRow( int index )
    {
      if( ( index < 0 ) || ( index > this.RowCount ) )
        throw new IndexOutOfRangeException();

      var content = new List<XElement>();
      for( int i = 0; i < this.ColumnCount; i++ )
      {
        var cell = ( this.GetColumnWidth( i ) != double.NaN )
                   ? HelperFunctions.CreateTableCell( this.GetColumnWidth( i ) * 20 )
                   : HelperFunctions.CreateTableCell();
        content.Add( cell );
      }

      return this.InsertRow( content, index );
    }

    public Row InsertRow( Row row, int index, bool keepFormatting = false )
    {
      if( row == null )
        throw new ArgumentNullException( "row" );

      if( index < 0 || index > RowCount )
        throw new IndexOutOfRangeException();

      List<XElement> content;
      if( keepFormatting )
        content = row.Xml.Elements().Select( element => HelperFunctions.CloneElement( element ) ).ToList();
      else
        content = row.Xml.Elements( XName.Get( "tc", Document.w.NamespaceName ) ).Select( element => HelperFunctions.CloneElement( element ) ).ToList();

      return InsertRow( content, index );
    }

    public void InsertColumn( int index, bool direction )
    {
      var colCount = this.ColumnCount;

      if( ( index < 0 ) || ( index >= colCount ) )
        throw new NullReferenceException( "index should be greater or equal to 0 and smaller to this.ColumnCount." );

      var rows = this.Rows;
      var firstRow = rows[ 0 ];

      var newColumnWidth = direction
                             ? ( index < firstRow.Cells.Count - 1 ) ? firstRow.Cells[ index + 1 ].Width : firstRow.Cells[ index ].Width
                             : firstRow.Cells[ index ].Width;

      if( this.RowCount > 0 )
      {
        _cachedColumnCount = -1;

        foreach( var row in rows )
        {
          // create cell (width will be set lower)
          var cell = HelperFunctions.CreateTableCell();

          // insert cell 
          if( row.Cells.Count < colCount )
          {
            int gridAfterValue = row.GridAfter;
            int currentPosition = 0;
            int posIndex = 0;

            foreach( var rowCell in row.Cells )
            {
              int gridSpanValue = ( rowCell.GridSpan != 0 ) ? rowCell.GridSpan - 1 : 0;

              // Check if the cell have a  gridSpan and if the index is between the lowest and highest cell value
              if( ( ( index - gridAfterValue ) >= currentPosition )
                && ( ( index - gridAfterValue ) <= ( currentPosition + gridSpanValue ) ) )
              {
                var dir = ( direction && ( index == ( currentPosition + gridSpanValue ) ) );
                this.AddCellToRow( row, cell, posIndex, dir );
                break;
              }

              ++posIndex;
              currentPosition += ( gridSpanValue + 1 );
            }
          }
          else
          {
            this.AddCellToRow( row, cell, index, direction );
          }
        }

        var newWidths = new List<float>( colCount + 1 );
        this.ColumnWidths.ForEach( pWidth => newWidths.Add( Convert.ToSingle( pWidth ) ) );
        newWidths.Insert( index, Convert.ToSingle( newColumnWidth ) );

        this.SetWidths( newWidths.ToArray(), false );

        this.Document.ClearParagraphsCache();
      }
    }

    public override void InsertPageBreakBeforeSelf()
    {
      base.InsertPageBreakBeforeSelf();
    }

    public void SetWidths( float[] widths, bool fixWidths = true )
    {
      if( widths == null )
        return;

      var totalTableWidth = widths.Sum();
      var availableWidth = this.Document.GetAvailableWidth();
      // Using autoFit and total columns size exceed page size => use a percentage of the page.

      if( ( this.AutoFit != AutoFit.Fixed ) && ( totalTableWidth > availableWidth ) )
      {
        var newWidths = new List<float>( widths.Length );
        widths.ToList().ForEach( pWidth =>
        {
          newWidths.Add( pWidth / totalTableWidth * 100 );
        } );
        this.SetWidthsPercentage( newWidths.ToArray(), Convert.ToSingle( availableWidth ) );
      }
      else
      {
        for( int i = 0; i < widths.Length; ++i )
        {
          this.SetColumnWidth( i, widths[ i ], fixWidths );
        }
      }
    }

    public void SetWidthsPercentage( float[] widthsPercentage, float? totalWidth = null )
    {
      if( totalWidth == null )
      {
        totalWidth = Convert.ToSingle( this.Document.GetAvailableWidth() );
      }

      List<float> widths = new List<float>( widthsPercentage.Length );
      widthsPercentage.ToList().ForEach( pWidth =>
      {
        widths.Add( ( pWidth * totalWidth.Value / 100 ) * ( 96 / 72 ) );
      } );
      //Case 173653 : Using SetColumnWidth instead of SetWidths() to update gridCol in XML.
      for( int i = 0; i < widths.Count; ++i )
      {
        this.SetColumnWidth( i, widths[ i ], false );
      }
    }

    public override void InsertPageBreakAfterSelf()
    {
      base.InsertPageBreakAfterSelf();
    }

    public override Table InsertTableBeforeSelf( Table t )
    {
      return base.InsertTableBeforeSelf( t );
    }

    public override Table InsertTableBeforeSelf( int rowCount, int columnCount )
    {
      return base.InsertTableBeforeSelf( rowCount, columnCount );
    }

    public override Table InsertTableAfterSelf( Table t )
    {
      return base.InsertTableAfterSelf( t );
    }

    public override Table InsertTableAfterSelf( int rowCount, int columnCount )
    {
      return base.InsertTableAfterSelf( rowCount, columnCount );
    }

    public override Paragraph InsertParagraphBeforeSelf( Paragraph p )
    {
      return base.InsertParagraphBeforeSelf( p );
    }

    public override Paragraph InsertParagraphBeforeSelf( string text )
    {
      return base.InsertParagraphBeforeSelf( text );
    }

    public override Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges )
    {
      return base.InsertParagraphBeforeSelf( text, trackChanges );
    }

    public override Paragraph InsertParagraphBeforeSelf( string text, bool trackChanges, Formatting formatting )
    {
      return base.InsertParagraphBeforeSelf( text, trackChanges, formatting );
    }

    public override Paragraph InsertParagraphAfterSelf( Paragraph p )
    {
      return base.InsertParagraphAfterSelf( p );
    }

    public override Paragraph InsertParagraphAfterSelf( string text, bool trackChanges, Formatting formatting )
    {
      return base.InsertParagraphAfterSelf( text, trackChanges, formatting );
    }

    public override Paragraph InsertParagraphAfterSelf( string text, bool trackChanges )
    {
      return base.InsertParagraphAfterSelf( text, trackChanges );
    }

    public override Paragraph InsertParagraphAfterSelf( string text )
    {
      return base.InsertParagraphAfterSelf( text );
    }

    public void SetBorder( TableBorderType borderType, Border border )
    {
      /*
       * Get the tblPr (table properties) element for this Table,
       * null will be return if no such element exists.
       */
      var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
      if( tblPr == null )
      {
        this.Xml.SetElementValue( XName.Get( "tblPr", Document.w.NamespaceName ), string.Empty );
        tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
      }

      /*
       * Get the tblBorders (table borders) element for this Table,
       * null will be return if no such element exists.
       */
      var tblBorders = tblPr.Element( XName.Get( "tblBorders", Document.w.NamespaceName ) );
      if( tblBorders == null )
      {
        tblPr.SetElementValue( XName.Get( "tblBorders", Document.w.NamespaceName ), string.Empty );
        tblBorders = tblPr.Element( XName.Get( "tblBorders", Document.w.NamespaceName ) );
      }

      /*
       * Get the 'borderType' (table border) element for this Table,
       * null will be return if no such element exists.
       */
      string tbordertype;
      tbordertype = borderType.ToString();
      // only lower the first char of string (because of insideH and insideV)
      tbordertype = tbordertype.Substring( 0, 1 ).ToLower() + tbordertype.Substring( 1 );

      var tblBorderType = tblBorders.Element( XName.Get( borderType.ToString(), Document.w.NamespaceName ) );
      if( tblBorderType == null )
      {
        tblBorders.SetElementValue( XName.Get( tbordertype, Document.w.NamespaceName ), string.Empty );
        tblBorderType = tblBorders.Element( XName.Get( tbordertype, Document.w.NamespaceName ) );
      }

      // get string value of border style
      var borderstyle = border.Tcbs.ToString().Substring( 5 );
      borderstyle = borderstyle.Substring( 0, 1 ).ToLower() + borderstyle.Substring( 1 );

      // The val attribute is used for the border style
      tblBorderType.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), borderstyle );

      if( border.Tcbs != BorderStyle.Tcbs_nil )
      {
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
        tblBorderType.SetAttributeValue( XName.Get( "sz", Document.w.NamespaceName ), ( size ).ToString() );

        // The space attribute is used for the cell spacing (probably '0')
        tblBorderType.SetAttributeValue( XName.Get( "space", Document.w.NamespaceName ), ( border.Space ).ToString() );

        // The color attribute is used for the border color
        tblBorderType.SetAttributeValue( XName.Get( "color", Document.w.NamespaceName ), border.Color.ToHex() );
      }
    }

    public Border GetBorder( TableBorderType borderType )
    {
      // instance with default border values
      var b = new Border();

      /*
       * Get the tblPr (table properties) element for this Table,
       * null will be return if no such element exists.
       */
      var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
      if( tblPr == null )
      {
        // uses default border style
        return b;
      }

      /*
       * Get the tblBorders (table borders) element for this Table,
       * null will be return if no such element exists.
       */
      var tblBorders = tblPr.Element( XName.Get( "tblBorders", Document.w.NamespaceName ) );
      if( tblBorders == null )
      {
        // uses default border style
        return b;
      }

      /*
       * Get the 'borderType' (table border) element for this Table,
       * null will be return if no such element exists.
       */
      string tbordertype = borderType.ToString();
      // only lower the first char of string (because of insideH and insideV)
      tbordertype = tbordertype.Substring( 0, 1 ).ToLower() + tbordertype.Substring( 1 );

      var tblBorderType = tblBorders.Element( XName.Get( tbordertype, Document.w.NamespaceName ) );
      if( tblBorderType == null )
      {
        // uses default border style
        return b;
      }

      // The val attribute is used for the border style
      var val = tblBorderType.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
      // If val is null, this table contains no border information.
      if( val == null )
      {
        // uses default border style
      }
      else
      {
        try
        {
          var bordertype = "Tcbs_" + val.Value;
          b.Tcbs = ( BorderStyle )Enum.Parse( typeof( BorderStyle ), bordertype );
        }
        catch
        {
          val.Remove();
          // uses default border style
        }
      }

      // The sz attribute is used for the border size
      var sz = tblBorderType.Attribute( XName.Get( "sz", Document.w.NamespaceName ) );
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
        {
          sz.Remove();
        }
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
      var space = tblBorderType.Attribute( XName.Get( "space", Document.w.NamespaceName ) );
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
      var color = tblBorderType.Attribute( XName.Get( "color", Document.w.NamespaceName ) );
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

    public Double GetColumnWidth( Int32 columnIndex )
    {
      List<Double> columnWidths = this.ColumnWidths;
      if( columnWidths == null || columnIndex > columnWidths.Count - 1 )
        return Double.NaN;

      return columnWidths[ columnIndex ];
    }

    public void SetColumnWidth( int columnIndex = 0, double width = -1, bool fixWidth = true )
    {
      var columnWidths = this.ColumnWidths;
      if( columnWidths == null || ( columnIndex > columnWidths.Count - 1 ) )
      {
        if( this.Rows.Count == 0 )
          throw new Exception( "There is at least one row required to detect the existing columns." );

        columnWidths = new List<Double>();
        var cells = this.Rows[ 0 ].Cells;
        foreach( var c in cells )
        {
          columnWidths.Add( c.Width );
        }

        // When some column are NaN, use a width based on the available page width.
        if( columnWidths.Contains( double.NaN ) )
        {
          var availablePageSpace = this.Document.GetAvailableWidth();
          var knownWidth = columnWidths.Where( c => !double.IsNaN( c ) );
          var columnSpaceUsed = knownWidth.Sum();
          var availableSpace = availablePageSpace - columnSpaceUsed;
          var unknownWidthColumnCount = columnWidths.Count - knownWidth.Count();
          var wantedColumnWidth = availableSpace / unknownWidthColumnCount;

          for( int i = 0; i < columnWidths.Count; ++i )
          {
            if( double.IsNaN( columnWidths[ i ] ) )
            {
              columnWidths[ i ] = wantedColumnWidth;
            }
          }
        }
      }

      // check if the columnIndex is valid 
      if( columnIndex > ( columnWidths.Count - 1 ) )
        throw new Exception( "The index is greather than the available table columns." );

      // append a new grid if null
      var grid = Xml.Element( XName.Get( "tblGrid", Document.w.NamespaceName ) );
      if( grid == null )
      {
        var tblPr = GetOrCreate_tblPr();
        tblPr.AddAfterSelf( new XElement( XName.Get( "tblGrid", Document.w.NamespaceName ) ) );
        grid = Xml.Element( XName.Get( "tblGrid", Document.w.NamespaceName ) );
      }

      // remove all existing values
      grid?.RemoveAll();

      // append new column widths
      int index = 0;
      foreach( var columnWidth in columnWidths )
      {
        double newWidth = columnWidth;
        if( ( index == columnIndex ) && ( width >= 0 ) )
        {
          newWidth = width;
        }

        var newColumn = new XElement( XName.Get( "gridCol", Document.w.NamespaceName ), new XAttribute( XName.Get( "w", Document.w.NamespaceName ), newWidth * 20d ) );
        grid?.Add( newColumn );
        index += 1;
      }

      // remove cell width
      if( width >= 0 )
      {
        foreach( var row in this.Rows )
        {
          if( columnIndex < row.Cells.Count )
          {
            row.Cells[ columnIndex ].Width = width;
          }
        }

        if( fixWidth )
        {
          // set AutoFit to Fixed
          this.AutoFit = AutoFit.Fixed;
        }
      }
    }

    public void SetTableCellMargin( TableCellMarginType type, double margin )
    {
      var tblPr = this.GetOrCreate_tblPr();
      var tblCellMarXName = XName.Get( "tblCellMar", Document.w.NamespaceName );
      var typeXName = XName.Get( type.ToString(), Document.w.NamespaceName );

      var tblCellMar = tblPr.Element( tblCellMarXName );
      if( tblCellMar == null )
      {
        tblPr.Add( new XElement( tblCellMarXName ) );
        tblCellMar = tblPr.Element( tblCellMarXName );
      }

      var side = tblCellMar.Element( typeXName );
      if( side == null )
      {
        tblCellMar.AddFirst( new XElement( typeXName ) );
        side = tblCellMar.Element( typeXName );
      }

      side.RemoveAttributes();
      // Set value and side for cell Margin
      side.Add( new XAttribute( XName.Get( "w", Document.w.NamespaceName ), margin * 20d ) );
      side.Add( new XAttribute( XName.Get( "type", Document.w.NamespaceName ), "dxa" ) );
    }

    public void DeleteAndShiftCellsLeft( int rowIndex, int celIndex )
    {
      var trPr = this.Rows[ rowIndex ].Xml.Element( XName.Get( "trPr", Document.w.NamespaceName ) );
      if( trPr != null )
      {
        var gridAfter = trPr.Element( XName.Get( "gridAfter", Document.w.NamespaceName ) );
        if( gridAfter != null )
        {
          var gridAfterValAttr = gridAfter.Attribute( XName.Get( "val", Document.w.NamespaceName ) );
          gridAfterValAttr.Value = ( gridAfterValAttr != null ) ? int.Parse( gridAfterValAttr.Value ).ToString() : "1";
        }
        else
        {
          gridAfter.SetAttributeValue( "val", 1 );
        }
      }
      else
      {
        var gridAfterXElement = new XElement( XName.Get( "gridAfter", Document.w.NamespaceName ) );
        var valXAttribute = new XAttribute( XName.Get( "val", Document.w.NamespaceName ), 1 );
        gridAfterXElement.Add( valXAttribute );

        var trPrXElement = new XElement( XName.Get( "trPr", Document.w.NamespaceName ) );
        trPrXElement.Add( gridAfterXElement );

        this.Rows[ rowIndex ].Xml.AddFirst( trPrXElement );
      }

      if( ( celIndex <= this.ColumnCount ) && ( this.Rows[ rowIndex ].ColumnCount <= this.ColumnCount ) )
      {
        this.Rows[ rowIndex ].Cells[ celIndex ].Xml.Remove();
      }
    }

    #endregion

    #region Internal Methods

    static internal AutoFit GetAutoFitFromXml( XElement xml )
    {
      if( xml == null )
        return AutoFit.ColumnWidth;

      // Get table attributes
      var tblWidthTypes = from d in xml.Descendants()
                          let type = d.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
                          where ( d.Name.LocalName == "tblW" ) && ( type != null )
                          select type;

      // Get column attributes
      var colWidthTypes = from d in xml.Descendants()
                          let type = d.Attribute( XName.Get( "type", Document.w.NamespaceName ) )
                          where ( d.Name.LocalName == "tcW" ) && ( type != null )
                          select type;

      if( tblWidthTypes.All( type => type.Value == "auto" ) && colWidthTypes.All( type => type.Value == "auto" ) )
        return AutoFit.Contents;
      if( tblWidthTypes.All( type => type.Value == "dxa" ) && colWidthTypes.All( type => type.Value == "dxa" ) )
        return AutoFit.Fixed;
      if( tblWidthTypes.All( type => type.Value == "pct" ) && colWidthTypes.All( type => type.Value == "pct" ) )
        return AutoFit.Window;

      return AutoFit.ColumnWidth;
    }

    internal XElement GetOrCreate_tblPr()
    {
      // Get the element.
      var tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );

      // If it dosen't exist, create it.
      if( tblPr == null )
      {
        this.Xml.AddFirst( new XElement( XName.Get( "tblPr", Document.w.NamespaceName ) ) );
        tblPr = Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
      }

      // Return the pPr element for this Paragraph.
      return tblPr;
    }

    internal void RemoveFirstRows( int rowCount )
    {
      Debug.Assert( rowCount >= 0 );

      var rowsXml = this.Xml.Elements( XName.Get( "tr", Document.w.NamespaceName ) );
      if( rowsXml != null )
      {
        rowsXml.Take( rowCount ).Remove();
      }
    }

    #endregion

    #region Private Methods

    private Row InsertRow( List<XElement> content, Int32 index )
    {
      var precedingRowGridAfter = 0;

      if( index > 0 )
      {
        precedingRowGridAfter = this.Rows[ index - 1 ].GridAfter;
        if( precedingRowGridAfter > 0 )
        {
          content.RemoveRange( content.Count - precedingRowGridAfter, precedingRowGridAfter );
        }
      }

      var newRow = new Row( this, Document, new XElement( XName.Get( "tr", Document.w.NamespaceName ), content ) );
      if( precedingRowGridAfter > 0 )
      {
        newRow.GridAfter = precedingRowGridAfter;
      }

      XElement rowXml;
      if( index == this.Rows.Count )
      {
        rowXml = this.Rows.Last().Xml;
        rowXml.AddAfterSelf( newRow.Xml );
      }

      else
      {
        rowXml = this.Rows[ index ].Xml;
        rowXml.AddBeforeSelf( newRow.Xml );
      }

      this.Document.ClearParagraphsCache();

      return newRow;
    }

    private void AddCellToRow( Row row, XElement cell, int index, bool direction )
    {
      if( index >= row.Cells.Count )
        throw new IndexOutOfRangeException( "index is greater or equals to row.Cells.Count." );

      if( direction )
      {
        row.Cells[ index ].Xml.AddAfterSelf( cell );
      }
      else
      {
        row.Cells[ index ].Xml.AddBeforeSelf( cell );
      }
    }

    private void UpdateTableLookXml()
    {
      var properties = this.Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
      Debug.Assert( properties != null, "properties shouldn't be null." );
      var tableLook = properties.Element( XName.Get( "tblLook", Document.w.NamespaceName ) );
      if( tableLook == null )
      {
        properties.Add( new XElement( XName.Get( "tblLook", Document.w.NamespaceName ) ) );
        tableLook = properties.Element( XName.Get( "tblLook", Document.w.NamespaceName ) );
      }
      // Using "val" is the old way of setting the tableLook.
      var val = tableLook.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );
      if( !string.IsNullOrEmpty( val ) )
      {
        var hexValue = 0x0;
        if( _tableLook.FirstRow )
        {
          hexValue += 0x20;
        }
        if( _tableLook.LastRow )
        {
          hexValue += 0x40;
        }
        if( _tableLook.FirstColumn )
        {
          hexValue += 0x80;
        }
        if( _tableLook.LastColumn )
        {
          hexValue += 0x100;
        }
        if( _tableLook.NoHorizontalBanding )
        {
          hexValue += 0x200;
        }
        if( _tableLook.NoVerticalBanding )
        {
          hexValue += 0x400;
        }
        tableLook.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), hexValue.ToString( "X4" ) );
      }
      else
      {
        tableLook.SetAttributeValue( XName.Get( "firstRow", Document.w.NamespaceName ), _tableLook.FirstRow ? "1" : "0" );
        tableLook.SetAttributeValue( XName.Get( "lastRow", Document.w.NamespaceName ), _tableLook.LastRow ? "1" : "0" );
        tableLook.SetAttributeValue( XName.Get( "firstColumn", Document.w.NamespaceName ), _tableLook.FirstColumn ? "1" : "0" );
        tableLook.SetAttributeValue( XName.Get( "lastColumn", Document.w.NamespaceName ), _tableLook.LastColumn ? "1" : "0" );
        tableLook.SetAttributeValue( XName.Get( "noHBand", Document.w.NamespaceName ), _tableLook.NoHorizontalBanding ? "1" : "0" );
        tableLook.SetAttributeValue( XName.Get( "noVBand", Document.w.NamespaceName ), _tableLook.NoVerticalBanding ? "1" : "0" );
      }
    }

    private void UpdateTableWidth( AutoFit autoFit, string widthValue )
    {
      switch( autoFit )
      {
        case AutoFit.Contents:
          if( this.Xml != null )
          {
            var tblW = from d in this.Xml.Descendants()
                       let type = d.Attribute( XName.Get( "w", Document.w.NamespaceName ) )
                       where ( d.Name.LocalName == "tblW" ) && ( type != null )
                       select type;

            foreach( var w in tblW )
            {
              w.Value = widthValue;
            }
          }
          break;

        case AutoFit.Window:
          if( this.Xml != null )
          {
            var tblPr = this.Xml.Descendants().FirstOrDefault( el => el.Name.LocalName == "tblPr" );
            if( tblPr == null )
            {
              this.Xml.AddFirst( new XElement( XName.Get( "tblPr", Document.w.NamespaceName ) ) );
              tblPr = this.Xml.Element( XName.Get( "tblPr", Document.w.NamespaceName ) );
            }

            // Set the table width to a percentage value.
            var tableWidths = tblPr.Descendants().Where( el => el.Name.LocalName == "tblW" );
            if( tableWidths != null )
            {
              tableWidths.Remove();
            }

            var tableWidth = new XElement( XName.Get( "tblW", Document.w.NamespaceName ),
                                           new XAttribute( XName.Get( "type", Document.w.NamespaceName ), "pct" ),
                                           new XAttribute( XName.Get( "w", Document.w.NamespaceName ), widthValue ) );

            tblPr.Add( tableWidth );
          }
          break;

        default:
          break;
      }
    }

    internal void RemoveInternal()
    {
      this.Xml.Remove();
    }

    #endregion

    #region Event Handlers

    private void TableLook_PropertyChanged( object sender, PropertyChangedEventArgs e )
    {
      this.UpdateTableLookXml();
    }

    #endregion




























  }

  public class Row : Container
  {
    #region Internal Members

    internal Table _table;

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
      trHeight.SetAttributeValue( XName.Get( "val", Document.w.NamespaceName ), ( ( int )( Math.Round( height * 20, 0 ) ) ).ToString( CultureInfo.InvariantCulture ) );
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

  public class Cell : Container
  {
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
    public Color Shading
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
          return Color.White;

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

    public Color FillColor
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
          return Color.Empty;

        int argb = Int32.Parse( fill.Value.Replace( "#", "" ), NumberStyles.HexNumber );
        return Color.FromArgb( argb );
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

  public class TableLook : INotifyPropertyChanged
  {
    #region Private members

    private bool _firstRow;
    private bool _lastRow;
    private bool _firstColumn;
    private bool _lastColumn;
    private bool _noHorizontalBanding;
    private bool _noVerticalBanding;

    #endregion

    #region Public Properties

    public bool FirstRow
    {
      get
      {
        return _firstRow;
      }
      set
      {
        _firstRow = value;
        OnPropertyChanged( "FirstRow" );
      }
    }

    public bool LastRow
    {
      get
      {
        return _lastRow;
      }
      set
      {
        _lastRow = value;
        OnPropertyChanged( "LastRow" );
      }
    }

    public bool FirstColumn
    {
      get
      {
        return _firstColumn;
      }
      set
      {
        _firstColumn = value;
        OnPropertyChanged( "FirstColumn" );
      }
    }

    public bool LastColumn
    {
      get
      {
        return _lastColumn;
      }
      set
      {
        _lastColumn = value;
        OnPropertyChanged( "LastColumn" );
      }
    }

    public bool NoHorizontalBanding
    {
      get
      {
        return _noHorizontalBanding;
      }
      set
      {
        _noHorizontalBanding = value;
        OnPropertyChanged( "NoHorizontalBanding" );
      }
    }

    public bool NoVerticalBanding
    {
      get
      {
        return _noVerticalBanding;
      }
      set
      {
        _noVerticalBanding = value;
        OnPropertyChanged( "NoVerticalBanding" );
      }
    }

    #endregion

    #region Constructors

    public TableLook()
    {
    }

    public TableLook( bool firstRow, bool lastRow, bool firstColumn, bool lastColumn, bool noHorizontalBanding, bool noVerticalBanding )
    {
      this.FirstRow = firstRow;
      this.LastRow = lastRow;
      this.FirstColumn = firstColumn;
      this.LastColumn = lastColumn;
      this.NoHorizontalBanding = noHorizontalBanding;
      this.NoVerticalBanding = noVerticalBanding;
    }

    #endregion

    #region INotifyPropertyChanged

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged( string propertyName )
    {
      if( PropertyChanged != null )
      {
        PropertyChanged( this, new PropertyChangedEventArgs( propertyName ) );
      }
    }

    #endregion
  }
}
