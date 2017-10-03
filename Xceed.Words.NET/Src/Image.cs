/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.IO.Packaging;
using System.IO;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Represents an Image embedded in a document.
  /// </summary>
  public class Image
  {
    #region Private Members

    /// <summary>
    /// A unique id which identifies this Image.
    /// </summary>
    private string _id;
    private DocX _document;

    #endregion

    #region Internal Members

    internal PackageRelationship _pr;

    #endregion

    #region Public Properties

    /// <summary>
    /// Returns the id of this Image.
    /// </summary>
    public string Id
    {
      get
      {
        return _id;
      }
    }

    ///<summary>
    /// Returns the name of the image file.
    ///</summary>
    public string FileName
    {
      get
      {
        return Path.GetFileName( _pr.TargetUri.ToString() );
      }
    }

    #endregion

    #region Constructors

    internal Image( DocX document, PackageRelationship pr )
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

    /// <summary>
    /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
    /// </summary>
    /// <returns></returns>
    /// <example>
    /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
    /// <code>
    /// using (DocX document = DocX.Create("Test.docx"))
    /// {
    ///    // Add an image to the document. 
    ///    Image     i = document.AddImage(@"Image.jpg");
    ///    
    ///    // Create a picture i.e. (A custom view of an image)
    ///    Picture   p = i.CreatePicture();
    ///    p.FlipHorizontal = true;
    ///    p.Rotation = 10;
    ///
    ///    // Create a new Paragraph.
    ///    Paragraph par = document.InsertParagraph();
    ///    
    ///    // Append content to the Paragraph.
    ///    par.Append("Here is a cool picture")
    ///       .AppendPicture(p)
    ///       .Append(" don't you think so?");
    ///
    ///    // Save all changes made to this document.
    ///    document.Save();
    /// }
    /// </code>
    /// </example>
    /// 
    public Picture CreatePicture()
    {
      return this.CreatePicture( -1, -1 );
    }

    /// <summary>
    /// Add an image to a document with specific height and width, create a custom view of that image (picture) and then insert it into a Paragraph using append.
    /// </summary>
    public Picture CreatePicture( int height, int width )
    {
      return Paragraph.CreatePicture( _document, _id, string.Empty, string.Empty, width, height );
    }

    #endregion
  }
}
