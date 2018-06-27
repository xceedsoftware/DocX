/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System.IO;
using System.Threading;

// This classe is used to prevent deadlocks. 
// It is based on https://stackoverflow.com/questions/21482820/openxml-hanging-while-writing-elements

namespace Xceed.Words.NET
{
  public class PackagePartStream : Stream
  {
    #region Private Members

    private static readonly object lockObject = new object();
    private readonly Stream _stream;

    #endregion

    #region constructors

    public PackagePartStream( Stream stream )
    {
      _stream = stream;
    }

    #endregion

    #region Overrides Properties

    public override bool CanRead
    {
      get
      {
        return _stream.CanRead;
      }
    }

    public override bool CanSeek
    {
      get
      {
        return _stream.CanSeek;
      }
    }

    public override bool CanWrite
    {
      get
      {
        return _stream.CanWrite;
      }
    }

    public override long Length
    {
      get
      {
        return _stream.Length;
      }
    }

    public override long Position
    {
      get
      {
        return _stream.Position;
      }

      set
      {
        _stream.Position = value;
      }
    }

    #endregion

    #region Overrides Methods

    public override long Seek( long offset, SeekOrigin origin )
    {
      return _stream.Seek( offset, origin );
    }

    public override void SetLength( long value )
    {
      _stream.SetLength( value );
    }

    public override int Read( byte[] buffer, int offset, int count )
    {
      return _stream.Read( buffer, offset, count );
    }

    public override void Write( byte[] buffer, int offset, int count )
    {
      lock(lockObject)
      {
        _stream.Write( buffer, offset, count );
      }
    }

    public override void Flush()
    {
      lock(lockObject)
      {
        _stream.Flush();
      }
    }

    public override void Close()
    {
      _stream.Close();
    }

    protected override void Dispose( bool disposing )
    {
      _stream.Dispose();
    }

    #endregion
  }
}
