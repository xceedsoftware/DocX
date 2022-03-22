/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.IO;

// This class is used to prevent deadlocks. 
// It is based on https://stackoverflow.com/questions/21482820/openxml-hanging-while-writing-elements

namespace Xceed.Document.NET
{
  internal class PackagePartStream : Stream
  {
    #region Private Members

    private static readonly object s_lockObject = new object();
    private readonly Stream m_stream;

    #endregion

    #region Constructors

    public PackagePartStream( Stream stream )
    {
      m_stream = stream;
    }

    #endregion

    #region Overrides Properties

    public override bool CanRead
    {
      get
      {
        return m_stream.CanRead;
      }
    }

    public override bool CanSeek
    {
      get
      {
        return m_stream.CanSeek;
      }
    }

    public override bool CanWrite
    {
      get
      {
        return m_stream.CanWrite;
      }
    }

    public override long Length
    {
      get
      {
        return m_stream.Length;
      }
    }

    public override long Position
    {
      get
      {
        return m_stream.Position;
      }

      set
      {
        m_stream.Position = value;
      }
    }

    #endregion

    #region Overrides Methods

    public override long Seek( long offset, SeekOrigin origin )
    {
      return m_stream.Seek( offset, origin );
    }

    public override void SetLength( long value )
    {
      m_stream.SetLength( value );
    }

    public override int Read( byte[] buffer, int offset, int count )
    {
      return m_stream.Read( buffer, offset, count );
    }

    public override void Write( byte[] buffer, int offset, int count )
    {
      lock( s_lockObject )
      {
        m_stream.Write( buffer, offset, count );
      }
    }

    public override void Flush()
    {
      lock( s_lockObject )
      {
        m_stream.Flush();
      }
    }

    public override void Close()
    {
      m_stream.Close();
    }

    protected override void Dispose( bool disposing )
    {
      m_stream.Dispose();
    }

    #endregion
  }
}
