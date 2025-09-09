/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public class LineSeries : Series
  {

    #region Constructors
    public LineSeries( string name ) : base( name )
    {
    }

    internal LineSeries( XElement xml ) : base( xml )
    {
    }

    #endregion // Constructors

    #region Overrides

    protected override XElement GetSpPrElement( XElement colorData, string widthValue = null )
    {
      if( string.IsNullOrEmpty( widthValue ) )
      {
        return new XElement( XName.Get( "spPr", Document.c.NamespaceName ),
            new XElement( XName.Get( "ln", Document.a.NamespaceName ), colorData ) );
      }
      else
      {
        return base.GetSpPrElement( colorData, widthValue );
      }
    }

    #endregion // Overrides
  }
}
