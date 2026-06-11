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
using System.IO.Packaging;
using System.IO;
using System.Globalization;
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
          _alignment = (Alignment)Enum.Parse( typeof( Alignment ), val.Value );
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
            this.Design = (TableDesign)Enum.Parse( typeof( TableDesign ), cleanValue );
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

      // If the original cell has borders, re-apply them in the merged cells.
      var start_tcBorders = start_tcPr.Element( XName.Get( "tcBorders", Document.w.NamespaceName ) );
      if( start_tcBorders != null )
      {
        // for all merged cells with the starting one...
        validRows = this.Rows.GetRange( startRow + 1, endRow - ( startRow + 1 ) + 1 );
        foreach( var row in validRows )
        {
          var c = row.Cells[ columnIndex ];
          var tcPr = c.Xml.Element( XName.Get( "tcPr", Document.w.NamespaceName ) );
          var tcBorders = tcPr.Element( XName.Get( "tcBorders", Document.w.NamespaceName ) );
          if( tcBorders != null )
          {
            tcBorders.Remove();
          }
          tcPr.Add( start_tcBorders );
        }
      }
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
          b.Tcbs = (BorderStyle)Enum.Parse( typeof( BorderStyle ), bordertype );
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

    private Cell FindTrueColCell( Row row, int trueColIndex )
    {

      int newCellTrueIndex = 0;

      for( int i = 0; i < row.Cells.Count; i++ )
      {
        XElement gridspanEl = row.Cells[ i ].Xml?.Descendants( Document.w + "gridSpan" )?.FirstOrDefault();

        if( gridspanEl != null && int.TryParse( gridspanEl?.Attribute( Document.w + "val" )?.Value, out int result ) )
        {
          newCellTrueIndex += result;
        }
        else
        {
          newCellTrueIndex++;
        }

        if( newCellTrueIndex > trueColIndex )
        {
          return row.Cells[ i ];
        }
      }

      return null;
    }

    private static void ModifyVMerge( XElement tc, string val )
    {
      var tcPr = tc.Element( Document.w + "tcPr" );
      var vMerge = tcPr.Element( Document.w + "vMerge" );

      if( vMerge == null )
      {
        tcPr.Add( new XElement( Document.w + "vMerge", new XAttribute( Document.w + "val", val ) ) );
      }
      else
      {
        tcPr.Element( Document.w + "vMerge" ).SetAttributeValue( Document.w + "val", val );
      }
    }

    private static void ArangeGridSpanOnNewCell( int splitSpan, XElement newElement )
    {
      if( splitSpan > 1 )
      {
        var gridSpan = newElement.Element( Document.w + "tcPr" )?.Element( Document.w + "gridSpan" );

        if( gridSpan != null )
        {
          gridSpan.SetAttributeValue( Document.w + "val", splitSpan.ToString() );
        }
        else
        {
          newElement.Element( Document.w + "tcPr" )?.Add( new XElement( Document.w + "gridSpan", new XAttribute( Document.w + "val", splitSpan.ToString() ) ) );
        }
      }
      else
      {
        var gridSpan = newElement.Element( Document.w + "tcPr" )?.Element( Document.w + "gridSpan" );
        if( gridSpan != null )
        {
          gridSpan.Remove();
        }
      }
    }


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
      if( this.Xml.Parent != null )
      {

        this.Xml.Remove();
      }
    }

    #endregion

    #region Event Handlers

    private void TableLook_PropertyChanged( object sender, PropertyChangedEventArgs e )
    {
      this.UpdateTableLookXml();
    }

    #endregion




























  }
}
