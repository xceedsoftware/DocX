/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.IO.Packaging;
using Xceed.Drawing;
using System.Xml.Linq;
using System.Globalization;
using System.Linq;
using System;

namespace Xceed.Document.NET
{
  public abstract class BaseSeries
  {

    #region Public Properties





    public virtual Color Color
    {
      get;
      set;
    }

    #endregion // Public Properties

    #region Internal Properties
    internal XElement Xml
    {
      get; private set;
    }

    internal virtual PackagePart PackagePart
    {
      get;
      set;
    }

    #endregion // Internal Properties

    #region Contructors

    internal BaseSeries( string name )
    {
    }

    internal BaseSeries( XElement xml )
    {
      this.SetXml( xml );
    }

    #endregion // Contructors

    #region Internal Method







    #endregion

    #region Protected Methods

    protected virtual void InitializeDataPoint()
    {
    }

    protected virtual void InsertDataPoint( DataPointBase dataPoint )
    {
    }

    protected void SetXml( XElement xml )
    {
      this.Xml = xml;
    }

    #endregion // Protected Methods
  }
}
