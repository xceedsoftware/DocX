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
using System.IO.Packaging;
using System.IO;
using System.Linq;

namespace Xceed.Document.NET
{
  public class Image
  {
    #region Private Members

    private string _id;
    private Document _document;

    #endregion

    #region Internal Members

    internal PackageRelationship _pr;

    #endregion

    #region Public Properties

    public string Id
    {
      get
      {
        return _id;
      }
    }

    public string FileName
    {
      get
      {
        return Path.GetFileName( _pr.TargetUri.ToString() );
      }
    }

    #endregion

    #region Constructors

    internal Image( Document document, PackageRelationship pr )
    {
      _document = document;
      _pr = pr;
      _id = pr.Id;
    }

    #endregion

    #region Public Methods

    public Stream GetStream( FileMode mode, FileAccess access )
    {
      string temp = _pr.SourceUri.OriginalString;
      string start = temp.Remove( temp.LastIndexOf( '/' ) );
      string end = _pr.TargetUri.OriginalString;
      string full = end.Contains( start ) ? end : start + "/" + end;

      return ( new PackagePartStream( _document._package.GetPart( new Uri( full, UriKind.Relative ) ).GetStream( mode, access ) ) );
    }

    public Picture CreatePicture()
    {
      return this.CreatePicture( -1f, -1f );
    }

    public Picture CreatePicture( float height, float width )
    {
      return Paragraph.CreatePicture( _document, _id, string.Empty, string.Empty, width, height );
    }

    public void Remove()
    {
      // No more of this image in the Document.
      if( !_document.Pictures.Any( picture => picture.Id == this.Id ) )
      {
        if( _pr.Package != null )
        {
          var uriString = _pr.TargetUri.OriginalString;
          if( !uriString.StartsWith( "/" ) )
          {
            uriString = "/" + uriString;
          }
          if( !uriString.StartsWith( "/word/" ) )
          {
            uriString = "/word" + uriString;
          }

          var uri = new Uri( uriString, UriKind.Relative );

          _pr.Package.DeletePart( uri );
        }

        if( _document.PackagePart != null )
        {
          _document.PackagePart.DeleteRelationship( _id );
        }
      }
    }

    #endregion
  }
}
