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
using System.IO;
using System.IO.Packaging;
using System.Linq;
#if NETFRAMEWORK
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using Microsoft.Win32;
using System.Drawing.Imaging;
using System.Xml;
#endif
using System.Xml.Linq;
using Xceed.Document.NET;

namespace Xceed.Words.NET
{
  public class DocX : Xceed.Document.NET.Document
  {
    private static bool IsCommercialLicenseRead = false;
    private bool _canClosePackage = true;

    #region Constructors

    internal DocX( Xceed.Document.NET.Document document, XElement xml )
        : base( document, xml )
    {
      if( !DocX.IsCommercialLicenseRead )
      {
        Console.WriteLine( "===================================================================\n"
                         + "Thank you for using Xceed's DocX library.\n"
                         + "Please note that this software is used for non-commercial use only.\n"
                         + "To obtain a commercial license, please visit www.xceed.com.\n"
                         + "===================================================================" );

        DocX.IsCommercialLicenseRead = true;
      }
    }

    #endregion

    #region Public Methods

    public static DocX Create( Stream stream, DocumentTypes documentType = DocumentTypes.Document )
    {
      var docX = new DocX( null, null ) as Xceed.Document.NET.Document;
      Xceed.Document.NET.Document.PrepareDocument( ref docX, documentType );
      docX._stream = stream;
      return docX as DocX;
    }

    public static DocX Create( string filename, DocumentTypes documentType = DocumentTypes.Document )
    {
      var docX = new DocX( null, null ) as Xceed.Document.NET.Document;
      Xceed.Document.NET.Document.PrepareDocument( ref docX, documentType );
      docX.SetFileName( filename );
      return docX as DocX;
    }

    public static DocX Load( Stream stream )
    {
      var docX = new DocX( null, null ) as Xceed.Document.NET.Document;
      return Xceed.Document.NET.Document.Load( stream, docX, DocumentTypes.Document ) as DocX;
    }

    public static DocX Load( string filename )
    {
      var docX = new DocX( null, null ) as Xceed.Document.NET.Document;
      return Xceed.Document.NET.Document.Load( filename, docX, DocumentTypes.Document ) as DocX;
    }


















































    #endregion

    #region Overrides

    public override void SaveAs( Stream stream, string password = "" )
    {
      if( this.IsPackageClosed( _package ) )
      {
        // When package is closed (already saved), reload the package and restart SaveAs();
        var initialDoc = DocX.ReloadDocument( this );
        initialDoc.SaveAs( stream, password );
        return;
      }

      base.SaveAs( stream, password );
    }

    public override void SaveAs( string filename, string password = "" )
    {
      if( this.IsPackageClosed( _package ) )
      {
        // When package is closed (already saved), reload the package and restart SaveAs();
        var initialDoc = DocX.ReloadDocumentFromFileName( this );
        initialDoc.SaveAs( filename, password );
        return;
      }

      base.SaveAs( filename, password );
    }

    public override void Save( string password = "" )
    {
      Message.ShowMessage();



      if( this.IsPackageClosed( _package ) )
      {
        // When package is closed (already saved), reload the package and restart Save();
        var initialDoc = DocX.ReloadDocumentFromFileName( this );
        initialDoc.Save();
        return;
      }

      if( ( _settings == null ) )
      {
        using( TextReader textReader = new StreamReader( _settingsPart.GetStream() ) )
        {
          _settings = XDocument.Load( textReader );
        }
      }

      ValidatePasswordProtection( password );

      // Save the main document
      using( TextWriter tw = new StreamWriter( new PackagePartStream( this.PackagePart.GetStream( FileMode.Create, FileAccess.Write ) ) ) )
      {
        _mainDoc.Save( tw, SaveOptions.None );
      }

      // Save the header/footer content.
      this.SaveHeadersFooters();

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

      if( _canClosePackage )
      {
        // Close the document so that all it's sub-part can be saved and it can be saved in .NETStandard/NET5.
        _package.Close();
      }

      #region Save this document back to a file or stream, that was specified by the user at save time.
      this.WriteToFileOrStream();
      #endregion
    }

    public override Xceed.Document.NET.Document Copy()
    {
      return this.InternalCopy();
    }













    #endregion

    #region Internal Methods

    protected internal override Xceed.Document.NET.Document InternalCopy( bool closePackage = false )
    {
      try
      {
        var initialDoc = this;
        if( this.IsPackageClosed( _package ) )
        {
          initialDoc = DocX.ReloadDocumentFromFileName( this ) as DocX;
        }

        initialDoc._canClosePackage = closePackage;

        initialDoc._isCopyingDocument = true;
        var memorystream = new MemoryStream();
        initialDoc.SaveAs( memorystream );
        initialDoc._isCopyingDocument = false;

        initialDoc._canClosePackage = true;

        memorystream.Seek( 0, SeekOrigin.Begin );
        var doc = DocX.Load( memorystream );
        doc.SetFileName( initialDoc._filename );

        return doc;
      }
      catch( Exception )
      {
        throw new InvalidOperationException( "The copy of the document could not be done." );
      }
    }

    protected internal override void SaveHeadersFooters()
    {
      foreach( var section in this.Sections )
      {
        var headers = section.Headers;
        var footers = section.Footers;

        // Header Even
        if( (headers.Even != null) && (headers.Even.Xml != null) && this.PackagePart.RelationshipExists( headers.Even.Id ) )
        {
          var target = PackUriHelper.ResolvePartUri
          (
              this.PackagePart.Uri,
              this.PackagePart.GetRelationship( headers.Even.Id ).TargetUri
          );
          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                headers.Even.Xml
            ).Save( tw, SaveOptions.None );
          }
        }

        // Header Odd
        if( (headers.Odd != null) && (headers.Odd.Xml != null) && this.PackagePart.RelationshipExists( headers.Odd.Id ) )
        {
          var target = PackUriHelper.ResolvePartUri
          (
             this.PackagePart.Uri,
             this.PackagePart.GetRelationship( headers.Odd.Id ).TargetUri
          );

          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                headers.Odd.Xml
            ).Save( tw, SaveOptions.None );
          }
        }

        // Header First
        if( (headers.First != null) && (headers.First.Xml != null) && this.PackagePart.RelationshipExists( headers.First.Id ) )
        {
          var target = PackUriHelper.ResolvePartUri
          (
            this.PackagePart.Uri,
            this.PackagePart.GetRelationship( headers.First.Id ).TargetUri
          );

          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                headers.First.Xml
            ).Save( tw, SaveOptions.None );
          }
        }

        // Footer Odd
        if( (footers.Odd != null) && (footers.Odd.Xml != null) && this.PackagePart.RelationshipExists( footers.Odd.Id ) )
        {
          var target = PackUriHelper.ResolvePartUri
          (
             this.PackagePart.Uri,
             this.PackagePart.GetRelationship( footers.Odd.Id ).TargetUri
          );

          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                footers.Odd.Xml
            ).Save( tw, SaveOptions.None );
          }
        }

        // Footer Even
        if( (footers.Even != null) && (footers.Even.Xml != null) && this.PackagePart.RelationshipExists( footers.Even.Id ) )
        {
          var target = PackUriHelper.ResolvePartUri
          (
            this.PackagePart.Uri,
            this.PackagePart.GetRelationship( footers.Even.Id ).TargetUri
          );

          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                footers.Even.Xml
            ).Save( tw, SaveOptions.None );
          }
        }

        // Footer First
        if( (footers.First != null) && (footers.First.Xml != null) && this.PackagePart.RelationshipExists( footers.First.Id ) )
        {
          var target = PackUriHelper.ResolvePartUri
          (
            this.PackagePart.Uri,
            this.PackagePart.GetRelationship( footers.First.Id ).TargetUri
          );

          using( TextWriter tw = new StreamWriter( new PackagePartStream( _package.GetPart( target ).GetStream( FileMode.Create, FileAccess.Write ) ) ) )
          {
            new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                footers.First.Xml
            ).Save( tw, SaveOptions.None );
          }
        }
      }
    }

    internal void ValidatePasswordProtection( string password )
    {
      if( !string.IsNullOrEmpty( password ) )
      {
        if( this.IsPasswordProtected )
        {
          if( _settings == null )
            throw new NullReferenceException();

          var documentProtection = _settings.Descendants( XName.Get( "documentProtection", w.NamespaceName ) ).FirstOrDefault();

          if( documentProtection != null )
          {
            var hash = documentProtection.Attribute( XName.Get( "hash", w.NamespaceName ) );

            if( hash != null )
            {
              var keyValues = ComputeHashValue( password, documentProtection );

              if( !string.IsNullOrEmpty( keyValues ) && hash.Value != keyValues )
              {
                throw new UnauthorizedAccessException( "Invalid password." );
              }
            }
          }
        }
      }
      else
      {
        if( this.IsPasswordProtected )
          throw new UnauthorizedAccessException( "The document is password protected, please set the document password in the Save()/SaveAs() method." );
      }
    }

    #endregion

    #region Private Method


























































    private void WriteToFileOrStream()
    {
      if( _filename != null && !_isCopyingDocument )
      {
        using( var fs = new FileStream( _filename, FileMode.Create ) )
        {
          if( _memoryStream.CanSeek )
          {
            // Write to the beginning of the stream
            _memoryStream.Position = 0;
            HelperFunctions.CopyStream( _memoryStream, fs );
          }
          else
            fs.Write( _memoryStream.ToArray(), 0, ( int )_memoryStream.Length );
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
    }

    private bool IsPackageClosed( Package package )
    {
      if( package == null )
        return true;

      try
      {
        var access = package.FileOpenAccess;
      }
      catch( Exception )
      {
        return true;
      }

      return false;
    }

    private static Xceed.Document.NET.Document ReloadDocument( Xceed.Document.NET.Document document )
    {
      var doc = ( ( document._stream != null ) && ( document._stream.Length > 0 ) ) ? DocX.Load( document._stream ) : DocX.Load( document._filename );
      doc.SetFileName( document._filename );

      return doc;
    }

    private static Xceed.Document.NET.Document ReloadDocumentFromFileName( Xceed.Document.NET.Document document )
    {
      var doc = ( !string.IsNullOrEmpty( document._filename ) && File.Exists( document._filename ) ) ? DocX.Load( document._filename ) : DocX.Load( document._stream );
      doc.SetFileName( document._filename );

      return doc;
    }

    #endregion
  }
}
