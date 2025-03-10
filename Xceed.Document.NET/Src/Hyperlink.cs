/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO.Packaging;

namespace Xceed.Document.NET
{
  public class Hyperlink : DocumentElement
  {
    #region Internal Members

    internal Uri uri;
    internal String text;

    internal Dictionary<PackagePart, PackageRelationship> hyperlink_rels;    
    internal int type;
    internal String id;
    internal XElement instrText;
    internal List<XElement> runs;

    #endregion

    #region Public Properties

    public string Text
    {
      get
      {
        return this.text;
      }

      set
      {
        XElement rPr =
            new XElement
            (
                Document.w + "rPr",
                new XElement
                (
                    Document.w + "rStyle",
                    new XAttribute( Document.w + "val", "Hyperlink" )
                )
            );

        // Format and add the new text.
        List<XElement> newRuns = HelperFunctions.FormatInput( value, rPr );

        if( type == 0 )
        {
          // Get all the runs in this Text.
          var runs = from r in Xml.Elements()
                     where r.Name.LocalName == "r"
                     select r;

          // Remove each run.
          for( int i = 0; i < runs.Count(); i++ )
            runs.Remove();

          Xml.Add( newRuns );
        }

        else
        {
          XElement separate = XElement.Parse( @"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='separate'/> 
                    </w:r>" );

          XElement end = XElement.Parse( @"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='end' /> 
                    </w:r>" );

          runs.Last().AddAfterSelf( separate, newRuns, end );
          runs.ForEach( r => r.Remove() );
        }

        this.text = value;
      }
    }

    public Uri Uri
    {
      get
      {
        if( (type == 0) && !String.IsNullOrEmpty(id) )
        {
          var r = this.PackagePart.GetRelationship( id );
          return r.TargetUri;
        }

        return this.uri;
      }

      set
      {
        if( type == 0 )
        {
          var r = this.PackagePart.GetRelationship( id );

          // Get all of the information about this relationship.
          var r_tm = r.TargetMode;
          var r_rt = r.RelationshipType;
          var r_id = r.Id;

          // Delete the relationship
          this.PackagePart.DeleteRelationship( r_id );
          this.PackagePart.CreateRelationship( value, r_tm, r_rt, r_id );
        }

        else
        {
          instrText.Value = "HYPERLINK " + "\"" + value + "\"";
        }

        this.uri = value;
      }
    }




    #endregion

    #region Internal Properties

    internal List<XElement> Runs
    {
      get
      {
        if( this.Xml == null )
          return null;

        var runsList =  from r in this.Xml.Elements()
                         where r.Name.LocalName == "r"
                         select r;
        return runsList.ToList();
      }
    }

    #endregion

    #region Constructors

    internal Hyperlink( Document document, PackagePart mainPart, XElement i ) : base( document, i )
    {
      this.type = 0;
      var idAttribute = i.Attribute( XName.Get( "id", Document.r.NamespaceName ) );
      if( idAttribute != null )
      {
        this.id = idAttribute.Value;
      }

      StringBuilder sb = new StringBuilder();
      HelperFunctions.GetTextRecursive( i, ref sb );
      this.text = sb.ToString();
      this.PackagePart = mainPart;
    }

    internal Hyperlink( Document document, XElement instrText, List<XElement> runs ) : base( document, null )
    {
      this.type = 1;
      this.instrText = instrText;
      this.runs = runs;

      int start = instrText.Value.IndexOf( "HYPERLINK \"" );
      if( start != -1 )
        start += "HYPERLINK \"".Length;
      int end = instrText.Value.IndexOf( "\"", Math.Max( 0, start ));
      if( start != -1 && end != -1 )
      {
        this.uri = new Uri( instrText.Value.Substring( start, end - start ), UriKind.Absolute );

        StringBuilder sb = new StringBuilder();
        HelperFunctions.GetTextRecursive( new XElement( XName.Get( "temp", Document.w.NamespaceName ), runs ), ref sb );
        this.text = sb.ToString();
      }
    }

    #endregion

    #region Public Methods

    public void Remove()
    {
      Xml.Remove();
    }

    #endregion
  }
}
