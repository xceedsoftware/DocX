/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.Linq;
using System.Xml.Linq;
using Xceed.Drawing;

namespace Xceed.Document.NET
{
  public abstract class CommonSeries : BaseSeries
  {
    #region Public Properties

    public override Color Color
    {
      get
      {
        var spPr = this.Xml.Element( XName.Get( "spPr", Document.c.NamespaceName ) );
        if( spPr == null )
          return Color.Transparent;

        var srgbClr = spPr.Descendants( XName.Get( "srgbClr", Document.a.NamespaceName ) ).FirstOrDefault();
        if( srgbClr != null )
        {
          var val = srgbClr.Attribute( XName.Get( "val" ) );
          if( val != null )
          {
            var rgb = Color.Parse( val.Value );
          }
        }

        return Color.Transparent;
      }
      set
      {
        var spPrElement = this.Xml.Element( XName.Get( "spPr", Document.c.NamespaceName ) );
        string widthValue = string.Empty;

        if( spPrElement != null )
        {
          var ln = spPrElement.Element( XName.Get( "ln", Document.a.NamespaceName ) );
          if( ln != null )
          {
            var val = ln.Attribute( XName.Get( "w" ) );
            if( val != null )
            {
              widthValue = val.Value;
            }
          }
          spPrElement.Remove();
        }

        var colorData = new XElement( XName.Get( "solidFill", Document.a.NamespaceName ),
                    new XElement( XName.Get( "srgbClr", Document.a.NamespaceName ),
                        new XAttribute( XName.Get( "val" ), value.ToHex() ) ) );

        spPrElement = this.GetSpPrElement( colorData, widthValue );
        this.Xml.Element( XName.Get( "tx", Document.c.NamespaceName ) ).AddAfterSelf( spPrElement );
      }
    }


    #endregion // Public Proerties

    #region Constructors

    protected CommonSeries( string name ) : base( name )
    {
    }

    internal CommonSeries( XElement xml ) : base( xml )
    {
      this.SetXml( xml );
    }

    #endregion // Constructors

    #region Protected Methods

    protected virtual XElement GetSpPrElement( XElement colorData, string widthValue = null )
    {
      XElement spPrElement;

      if( string.IsNullOrEmpty( widthValue ) )
      {
        spPrElement = new XElement( XName.Get( "spPr", Document.c.NamespaceName ), colorData );
      }
      else
      {
        spPrElement = new XElement( XName.Get( "spPr", Document.c.NamespaceName ),
            new XElement( XName.Get( "ln", Document.a.NamespaceName ),
                new XAttribute( XName.Get( "w" ), widthValue ),
                colorData ) );
      }

      return spPrElement;
    }

    #endregion // Protected Methods
  }
}
