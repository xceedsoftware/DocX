﻿/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2017 Xceed Software Inc.
 
   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using Xceed.Document.NET;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Represents a DocX document.
  /// </summary>
  public class DocX : Xceed.Document.NET.Document
  {
    #region Constructors

    internal DocX( Xceed.Document.NET.Document document, XElement xml )
        : base( document, xml )
    {
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Creates a document using a Stream.
    /// </summary>
    /// <param name="stream">The Stream to create the document from.</param>
    /// <param name="documentType"></param>
    /// <returns>Returns a Document object which represents the document.</returns>
    /// <example>
    /// Creating a document from a FileStream.
    /// <code>
    /// // Use a FileStream fs to create a new document.
    /// using(FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Create))
    /// {
    ///     // Load the document using fs
    ///     using (var document = DocX.Create(fs))
    ///     {
    ///         // Do something with the document here.
    ///
    ///         // Save all changes made to this document.
    ///         document.Save();
    ///     }// Release this document from memory.
    /// }
    /// </code>
    /// </example>
    /// <example>
    /// Creating a document in a SharePoint site.
    /// <code>
    /// using(SPSite mySite = new SPSite("http://server/sites/site"))
    /// {
    ///     // Open a connection to the SharePoint site
    ///     using(SPWeb myWeb = mySite.OpenWeb())
    ///     {
    ///         // Create a MemoryStream ms.
    ///         using (MemoryStream ms = new MemoryStream())
    ///         {
    ///             // Create a document using ms.
    ///             using (var document = DocX.Create(ms))
    ///             {
    ///                 // Do something with the document here.
    ///
    ///                 // Save all changes made to this document.
    ///                 document.Save();
    ///             }// Release this document from memory
    ///
    ///             // Add the document to the SharePoint site
    ///             web.Files.Add("filename", ms.ToArray(), true);
    ///         }
    ///     }
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Load(string)"/>
    /// <seealso cref="DocX.Save()"/>
    public static DocX Create( Stream stream, DocumentTypes documentType = DocumentTypes.Document )
    {
      var docX = new DocX( null, null ) as Xceed.Document.NET.Document;
      Xceed.Document.NET.Document.PrepareDocument( ref docX, documentType );
      docX._stream = stream;
      return docX as DocX;
    }

    /// <summary>
    /// Creates a document using a fully qualified or relative filename.
    /// </summary>
    /// <param name="filename">The fully qualified or relative filename.</param>
    /// <param name="documentType"></param>
    /// <returns>Returns a Document object which represents the document.</returns>
    /// <example>
    /// <code>
    /// // Create a document using a relative filename.
    /// using (var document = DocX.Create(@"..\Test.docx"))
    /// {
    ///     // Do something with the document here.
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory
    /// </code>
    /// <code>
    /// // Create a document using a relative filename.
    /// using (var document = DocX.Create(@"..\Test.docx"))
    /// {
    ///     // Do something with the document here.
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory
    /// </code>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Load(string)"/>
    /// <seealso cref="DocX.Save()"/>
    /// </example>
    public static DocX Create( string filename, DocumentTypes documentType = DocumentTypes.Document )
    {
      var docX = new DocX( null, null ) as Xceed.Document.NET.Document;
      Xceed.Document.NET.Document.PrepareDocument( ref docX, documentType );
      docX._filename = filename;
      return docX as DocX;
    }

    /// <summary>
    /// Loads a document into a Document object using a Stream.
    /// </summary>
    /// <param name="stream">The Stream to load the document from.</param>
    /// <returns>
    /// Returns a Document object which represents the document.
    /// </returns>
    /// <example>
    /// Loading a document from a FileStream.
    /// <code>
    /// // Open a FileStream fs to a document.
    /// using (FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Open))
    /// {
    ///     // Load the document using fs.
    ///     using (var document = DocX.Load(fs))
    ///     {
    ///         // Do something with the document here.
    ///            
    ///         // Save all changes made to the document.
    ///         document.Save();
    ///     }// Release this document from memory.
    /// }
    /// </code>
    /// </example>
    /// <example>
    /// Loading a document from a SharePoint site.
    /// <code>
    /// // Get the SharePoint site that you want to access.
    /// using (SPSite mySite = new SPSite("http://server/sites/site"))
    /// {
    ///     // Open a connection to the SharePoint site
    ///     using (SPWeb myWeb = mySite.OpenWeb())
    ///     {
    ///         // Grab a document stored on this site.
    ///         SPFile file = web.GetFile("Source_Folder_Name/Source_File");
    ///
    ///         // Document.Load requires a Stream, so open a Stream to this document.
    ///         Stream str = new MemoryStream(file.OpenBinary());
    ///
    ///         // Load the file using the Stream str.
    ///         using (var document = DocX.Load(str))
    ///         {
    ///             // Do something with the document here.
    ///
    ///             // Save all changes made to the document.
    ///             document.Save();
    ///         }// Release this document from memory.
    ///     }
    /// }
    /// </code>
    /// </example>
    /// <seealso cref="DocX.Load(string)"/>
    /// <seealso cref="DocX.Save()"/>
    public static DocX Load( Stream stream )
    {
      var docX = new DocX( null, null ) as Xceed.Document.NET.Document;
      return Xceed.Document.NET.Document.Load( stream, docX, DocumentTypes.Document ) as DocX;
    }

    /// <summary>
    /// Loads a document into a Document object using a fully qualified or relative filename.
    /// </summary>
    /// <param name="filename">The fully qualified or relative filename.</param>
    /// <returns>
    /// Returns a DocX object which represents the document.
    /// </returns>
    /// <example>
    /// <code>
    /// // Load a document using its fully qualified filename
    /// using (var document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Do something with the document here
    ///
    ///     // Save all changes made to document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// <code>
    /// // Load a document using its relative filename.
    /// using(var document = DocX.Load(@"..\..\Test.docx"))
    /// { 
    ///     // Do something with the document here.
    ///                
    ///     // Save all changes made to document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public static DocX Load( string filename )
    {
      var docX = new DocX( null, null ) as Xceed.Document.NET.Document;
      return Xceed.Document.NET.Document.Load( filename, docX, DocumentTypes.Document ) as DocX;
    }

    #endregion

    #region Overrides

    /// <summary>
    /// Save this document back to the location it was loaded from.
    /// </summary>
    /// <example>
    /// <code>
    /// // Load a document.
    /// using (var document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Add an Image from a file.
    ///     document.AddImage(@"C:\Example\Image.jpg");
    ///
    ///     // Save all changes made to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <seealso cref="DocX.Load(System.IO.Stream)"/>
    /// <seealso cref="DocX.Load(string)"/> 
    /// <!-- 
    /// Bug found and fixed by krugs525 on August 12 2009.
    /// Use TFS compare to see exact code change.
    /// -->
    public override void Save()
    {
      var headers = this.Headers;
      var footers = this.Footers;

      // Save the main document
      using( TextWriter tw = new StreamWriter( new PackagePartStream( this.PackagePart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        _mainDoc.Save( tw, SaveOptions.None );
      }

      if( ( _settings == null ) || !this.isProtected )
      {
        using( TextReader textReader = new StreamReader( _settingsPart.GetStream() ) )
        {
          _settings = XDocument.Load( textReader );
        }
      }

      var body = _mainDoc.Root.Element( w + "body" );
      var sectPr = body.Descendants( w + "sectPr" ).FirstOrDefault();

      if( sectPr != null )
      {
        var evenHeaderRef =
        (
            from e in _mainDoc.Descendants( w + "headerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "even", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( evenHeaderRef != null )
        {
          var even = headers.Even.Xml;

          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( evenHeaderRef ).TargetUri
          );

          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                even
            ).Save( tw, SaveOptions.None );
          }
        }

        var oddHeaderRef =
        (
            from e in _mainDoc.Descendants( w + "headerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "default", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( oddHeaderRef != null )
        {
          var odd = headers.Odd.Xml;

          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( oddHeaderRef ).TargetUri
          );

          // Save header1
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                odd
            ).Save( tw, SaveOptions.None );
          }
        }

        var firstHeaderRef =
        (
            from e in _mainDoc.Descendants( w + "headerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "first", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( firstHeaderRef != null )
        {
          var first = headers.First.Xml;
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( firstHeaderRef ).TargetUri
          );

          // Save header3
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                first
            ).Save( tw, SaveOptions.None );
          }
        }

        var oddFooterRef =
        (
            from e in _mainDoc.Descendants( w + "footerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "default", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( oddFooterRef != null )
        {
          var odd = footers.Odd.Xml;
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( oddFooterRef ).TargetUri
          );

          // Save header1
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                odd
            ).Save( tw, SaveOptions.None );
          }
        }

        var evenFooterRef =
        (
            from e in _mainDoc.Descendants( w + "footerReference" )
            let type = e.Attribute( w + "type" )
            where type != null && type.Value.Equals( "even", StringComparison.CurrentCultureIgnoreCase )
            select e.Attribute( r + "id" ).Value
         ).LastOrDefault();

        if( evenFooterRef != null )
        {
          var even = footers.Even.Xml;
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( evenFooterRef ).TargetUri
          );

          // Save header2
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                even
            ).Save( tw, SaveOptions.None );
          }
        }

        var firstFooterRef =
        (
             from e in _mainDoc.Descendants( w + "footerReference" )
             let type = e.Attribute( w + "type" )
             where type != null && type.Value.Equals( "first", StringComparison.CurrentCultureIgnoreCase )
             select e.Attribute( r + "id" ).Value
        ).LastOrDefault();

        if( firstFooterRef != null )
        {
          var first = footers.First.Xml;
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( firstFooterRef ).TargetUri
          );

          // Save header3
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                first
            ).Save( tw, SaveOptions.None );
          }
        }

        // Save the settings document.
        using( TextWriter tw = new StreamWriter( new PackagePartStream( _settingsPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
        {
          _settings.Save( tw, SaveOptions.None );
        }

        if( _endnotesPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _endnotesPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _endnotes.Save( tw, SaveOptions.None );
          }
        }

        if( _footnotesPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _footnotesPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _footnotes.Save( tw, SaveOptions.None );
          }
        }

        if( _stylesPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _stylesPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _styles.Save( tw, SaveOptions.None );
          }
        }

        if( _stylesWithEffectsPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _stylesWithEffectsPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _stylesWithEffects.Save( tw, SaveOptions.None );
          }
        }

        if( _numberingPart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _numberingPart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _numbering.Save( tw, SaveOptions.None );
          }
        }

        if( _fontTablePart != null )
        {
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _fontTablePart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            _fontTable.Save( tw, SaveOptions.None );
          }
        }
      }

      // Close the document so that it can be saved.
      _package.Flush();
      _package.Close();

      #region Save this document back to a file or stream, that was specified by the user at save time.
      if( _filename != null )
      {
        using( FileStream fs = new FileStream( _filename, FileMode.Create ) )
        {
          if( _memoryStream.CanSeek )
          {
            // Write to the beginning of the stream
            _memoryStream.Position = 0;
            HelperFunctions.CopyStream( _memoryStream, fs );
          }
          else
            fs.Write( _memoryStream.ToArray(), 0, (int)_memoryStream.Length );
        }
      }
      else if( _stream.CanSeek )  //Check if stream can be seeked to support System.Web.HttpResponseStream.
      {
        // Set the length of this stream to 0
        _stream.SetLength( 0 );

        // Write to the beginning of the stream
        _stream.Position = 0;

        _memoryStream.WriteTo( _stream );
        _memoryStream.Flush();
      }
      #endregion
    }

    /// <summary>
    /// Copy the Document into a new Document
    /// </summary>
    /// <returns>Returns a copy of a the Document</returns>
    public override Xceed.Document.NET.Document Copy()
    {
      var memorystream = new MemoryStream();
      this.SaveAs( memorystream );
      memorystream.Seek( 0, SeekOrigin.Begin );
      return DocX.Load( memorystream );
    }

    #endregion
  }
}
