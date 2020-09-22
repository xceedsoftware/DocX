/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using System.IO;
using System.Collections.ObjectModel;

namespace Xceed.Document.NET
{
  public abstract class Container : DocumentElement
  {
    #region Public Properties

    /// <summary>
    /// Returns a list of all Paragraphs inside this container.
    /// </summary>
    /// <example>
    /// <code>
    ///  Load a document.
    /// using (var document = DocX.Load(@"Test.docx"))
    /// {
    ///    // All Paragraphs in this document.
    ///    <![CDATA[List<Paragraph>]]> documentParagraphs = document.Paragraphs;
    ///    
    ///    // Make sure this document contains at least one Table.
    ///    if (document.Tables.Count() > 0)
    ///    {
    ///        // Get the first Table in this document.
    ///        Table t = document.Tables[0];
    ///
    ///        // All Paragraphs in this Table.
    ///        <![CDATA[List<Paragraph>]]> tableParagraphs = t.Paragraphs;
    ///    
    ///        // Make sure this Table contains at least one Row.
    ///        if (t.Rows.Count() > 0)
    ///        {
    ///            // Get the first Row in this document.
    ///            Row r = t.Rows[0];
    ///
    ///            // All Paragraphs in this Row.
    ///             <![CDATA[List<Paragraph>]]> rowParagraphs = r.Paragraphs;
    ///
    ///            // Make sure this Row contains at least one Cell.
    ///            if (r.Cells.Count() > 0)
    ///            {
    ///                // Get the first Cell in this document.
    ///                Cell c = r.Cells[0];
    ///
    ///                // All Paragraphs in this Cell.
    ///                <![CDATA[List<Paragraph>]]> cellParagraphs = c.Paragraphs;
    ///            }
    ///        }
    ///    }
    ///
    ///    // Save all changes to this document.
    ///    document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public virtual ReadOnlyCollection<Paragraph> Paragraphs
    {
      get
      {
        var paragraphs = this.GetParagraphs();
        this.InitParagraphs( paragraphs );

        return paragraphs.AsReadOnly();
      }
    }

    public virtual ReadOnlyCollection<Paragraph> ParagraphsDeepSearch
    {
      get
      {
        return this.Paragraphs;
        //var paragraphs = this.GetParagraphs( true );
        //this.InitParagraphs( paragraphs );

        //return paragraphs.AsReadOnly();
      }
    }

    public virtual IList<Section> Sections
    {
      get
      {
        return this.GetSections();
      }
    }

    public virtual List<Table> Tables
    {
      get
      {
        List<Table> tables =
        (
            from t in Xml.Descendants( Document.w + "tbl" )
            select new Table( Document, t, this.PackagePart )
        ).ToList();

        return tables;
      }
    }

    public virtual List<Hyperlink> Hyperlinks
    {
      get
      {
        List<Hyperlink> hyperlinks = new List<Hyperlink>();

        foreach( Paragraph p in Paragraphs )
          hyperlinks.AddRange( p.Hyperlinks );

        return hyperlinks;
      }
    }

    public virtual List<Picture> Pictures
    {
      get
      {
        List<Picture> pictures = new List<Picture>();

        foreach( Paragraph p in Paragraphs )
          pictures.AddRange( p.Pictures );

        return pictures;
      }
    }







    public virtual List<List> Lists
    {
      get
      {
        var lists = new List<List>();
        var list = new List( Document, Xml );

        foreach( var paragraph in Paragraphs )
        {
          if( paragraph.IsListItem )
          {
            if( list.CanAddListItem( paragraph ) )
            {
              list.AddItem( paragraph );
            }
            else
            {
              lists.Add( list );
              list = new List( Document, Xml );
              list.AddItem( paragraph );
            }
          }
        }
        lists.Add( list );
        return lists;
      }
    }

    #endregion

    #region Public Methods

    public IList<Section> GetSections()
    {
      var paragraphs = this.Paragraphs;
      var sections = new List<Section>();
      var sectionParagraphs = new List<Paragraph>();

      foreach( var paragraph in paragraphs )
      {
        var sectionInPara = paragraph.Xml.Descendants().FirstOrDefault( s => s.Name.LocalName == "sectPr" );

        if( sectionInPara != null )
        {
          sectionParagraphs.Add( paragraph );

          var section = new Section( Document, sectionInPara, sections.Count() > 0 ? sections.Select(s => s.Xml) : null );
          section.SectionParagraphs = sectionParagraphs;

          sections.Add( section );
          sectionParagraphs = new List<Paragraph>();
        }
        else
        {
          sectionParagraphs.Add( paragraph );
        }
      }

      var body = Xml.DescendantsAndSelf( XName.Get( "body", Document.w.NamespaceName ) ).FirstOrDefault();
      if( body != null )
      {
        var sectPrList = body.Elements( XName.Get( "sectPr", Document.w.NamespaceName ) );

        var baseSectionXml = sectPrList.LastOrDefault();
        if( baseSectionXml != null )
        {
          var baseSection = ( sections.Count > 0 ) ? new Section( Document, baseSectionXml, sections.Select( s => s.Xml ) )
                                                   : new Section( Document, baseSectionXml, ( sectPrList.Count() > 1 ) ? sectPrList.Take( sectPrList.Count() - 1 ) : null );
          baseSection.SectionParagraphs = sectionParagraphs;
          sections.Add( baseSection );
        }
      }


      return sections;
    }

    /// <summary>
    /// Sets the Direction of content.
    /// </summary>
    /// <param name="direction">Direction either LeftToRight or RightToLeft</param>
    /// <example>
    /// Set the Direction of content in a Paragraph to RightToLeft.
    /// <code>
    /// // Load a document.
    /// using (Document document = DocX.Load(@"Test.docx"))
    /// {
    ///    // Get the first Paragraph from this document.
    ///    Paragraph p = document.InsertParagraph();
    ///
    ///    // Set the Direction of this Paragraph.
    ///    p.Direction = Direction.RightToLeft;
    ///
    ///    // Make sure the document contains at lest one Table.
    ///    if (document.Tables.Count() > 0)
    ///    {
    ///        // Get the first Table from this document.
    ///        Table t = document.Tables[0];
    ///
    ///        /* 
    ///         * Set the direction of the entire Table.
    ///         * Note: The same function is available at the Row and Cell level.
    ///         */
    ///        t.SetDirection(Direction.RightToLeft);
    ///    }
    ///
    ///    // Save all changes to this document.
    ///    document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    public virtual void SetDirection( Direction direction )
    {
      foreach( Paragraph p in Paragraphs )
        p.Direction = direction;
    }

    public virtual List<int> FindAll( string str )
    {
      return FindAll( str, RegexOptions.None );
    }

    public virtual List<int> FindAll( string str, RegexOptions options )
    {
      List<int> list = new List<int>();

      foreach( Paragraph p in Paragraphs )
      {
        List<int> indexes = p.FindAll( str, options );

        for( int i = 0; i < indexes.Count(); i++ )
          indexes[ i ] += p._startIndex;

        list.AddRange( indexes );
      }

      return list;
    }

    /// <summary>
    /// Find all unique instances of the given Regex Pattern,
    /// returning the list of the unique strings found
    /// </summary>
    /// <param name="pattern"></param>
    /// <param name="options"></param>
    /// <returns></returns>
    public virtual List<string> FindUniqueByPattern( string pattern, RegexOptions options )
    {
      var rawResults = new List<string>();

      foreach( Paragraph p in Paragraphs )
      {   // accumulate the search results from all paragraphs
        var partials = p.FindAllByPattern( pattern, options );
        rawResults.AddRange( partials );
      }

      // this dictionary is used to collect results and test for uniqueness
      var uniqueResults = new Dictionary<string, int>();

      foreach( string currValue in rawResults )
      {
        if( !uniqueResults.ContainsKey( currValue ) )
        {   // if the dictionary doesn't have it, add it
          uniqueResults.Add( currValue, 0 );
        }
      }

      return uniqueResults.Keys.ToList();  // return the unique list of results
    }

    public virtual void ReplaceText( string searchValue,
                                     string newValue,
                                     bool trackChanges = false,
                                     RegexOptions options = RegexOptions.None,
                                     Formatting newFormatting = null,
                                     Formatting matchFormatting = null,
                                     MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch, 
                                     bool escapeRegEx = true, 
                                     bool useRegExSubstitutions = false,
                                     bool removeEmptyParagraph = true )
    {
      if( string.IsNullOrEmpty( searchValue ) )
      {
        throw new ArgumentException( "searchValue cannot be null or empty.", "searchValue" );
      }
      if( newValue == null )
      {
        throw new ArgumentException( "newValue cannot be null.", "newValue" );
      }

      foreach( var p in this.Paragraphs )
      {
        p.ReplaceText( searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, fo, escapeRegEx, useRegExSubstitutions, removeEmptyParagraph );
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="searchValue">The value to find.</param>
    /// <param name="regexMatchHandler">A Func who accepts the matching regex search group value and passes it to this to return the replacement string.</param>
    /// <param name="trackChanges">Enable or disable the track changes.</param>
    /// <param name="options">The Regex options.</param>
    /// <param name="newFormatting"></param>
    /// <param name="matchFormatting"></param>
    /// <param name="fo"></param>
    /// <param name="removeEmptyParagraph">Remove empty paragraph</param>
    public virtual void ReplaceText( string searchValue, 
                                     Func<string,string> regexMatchHandler, 
                                     bool trackChanges = false,
                                     RegexOptions options = RegexOptions.None,
                                     Formatting newFormatting = null,
                                     Formatting matchFormatting = null,
                                     MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch, 
                                     bool removeEmptyParagraph = true )
    {
      if( string.IsNullOrEmpty( searchValue ) )
      {
        throw new ArgumentException( "searchValue cannot be null or empty.", "searchValue" );
      }
      if( regexMatchHandler == null )
      {
        throw new ArgumentException( "regexMatchHandler cannot be null", "regexMatchHandler" );
      }

      foreach( var p in this.Paragraphs )
      {
        p.ReplaceText( searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, fo, removeEmptyParagraph );
      }
    }

    public virtual void ReplaceTextWithObject( string searchValue,
                                    DocumentElement objectToAdd,
                                    bool trackChanges = false,
                                    RegexOptions options = RegexOptions.None,
                                    Formatting matchFormatting = null,
                                    MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch,
                                    bool escapeRegEx = true,
                                    bool removeEmptyParagraph = true )
    {
      if( string.IsNullOrEmpty( searchValue ) )
      {
        throw new ArgumentException( "searchValue cannot be null or empty.", "searchValue" );
      }
      if( objectToAdd == null )
      {
        throw new ArgumentException( "objectToAdd cannot be null.", "objectToAdd" );
      }

      // ReplaceText in the container
      foreach( Paragraph p in this.Paragraphs )
      {
        p.ReplaceTextWithObject( searchValue, objectToAdd, trackChanges, options, matchFormatting, fo, escapeRegEx, removeEmptyParagraph );
      }
    }

    public virtual void InsertAtBookmark( string toInsert, string bookmarkName, Formatting formatting = null )
    {
      if( string.IsNullOrWhiteSpace( bookmarkName ) )
        throw new ArgumentException( "bookmark cannot be null or empty", "bookmarkName" );

      foreach( var paragraph in Paragraphs )
      {
        paragraph.InsertAtBookmark( toInsert, bookmarkName, formatting );
      }
    }

    public virtual Paragraph InsertParagraph( int index, string text, bool trackChanges )
    {
      return InsertParagraph( index, text, trackChanges, null );
    }

    public virtual Paragraph InsertParagraph()
    {
      return InsertParagraph( string.Empty, false );
    }

    public virtual Paragraph InsertParagraph( int index, Paragraph p )
    {
      var newXElement = new XElement( p.Xml );
      p.Xml = newXElement;

      var paragraph = HelperFunctions.GetFirstParagraphEffectedByInsert( Document, index );

      if( paragraph == null )
      {
        this.AddElementInXml( p.Xml );
      }
      else
      {
        var split = HelperFunctions.SplitParagraph( paragraph, index - paragraph._startIndex );

        paragraph.Xml.ReplaceWith
        (
            split[ 0 ],
            newXElement,
            split[ 1 ]
        );
      }
      this.SetParentContainer( p );
      return p;
    }

    public virtual Paragraph InsertParagraph( Paragraph p )
    {
      #region Styles
      XDocument style_document;

      if( p._styles.Count() > 0 )
      {
        var style_package_uri = new Uri( "/word/styles.xml", UriKind.Relative );
        if( !Document._package.PartExists( style_package_uri ) )
        {
          var style_package = Document._package.CreatePart( style_package_uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum );
          using( TextWriter tw = new StreamWriter( new PackagePartStream( style_package.GetStream() ) ) )
          {
            style_document = new XDocument
            (
                new XDeclaration( "1.0", "UTF-8", "yes" ),
                new XElement( XName.Get( "styles", Document.w.NamespaceName ) )
            );

            style_document.Save( tw );
          }
        }

        var styles_document = Document._package.GetPart( style_package_uri );
        using( TextReader tr = new StreamReader( styles_document.GetStream() ) )
        {
          style_document = XDocument.Load( tr );
          var styles_element = style_document.Element( XName.Get( "styles", Document.w.NamespaceName ) );

          var ids = from d in styles_element.Descendants( XName.Get( "style", Document.w.NamespaceName ) )
                    let a = d.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) )
                    where a != null
                    select a.Value;

          foreach( XElement style in p._styles )
          {
            // If styles_element does not contain this element, then add it.

            if( !ids.Contains( style.Attribute( XName.Get( "styleId", Document.w.NamespaceName ) ).Value ) )
            {
              styles_element.Add( style );
            }
          }
        }

        using( TextWriter tw = new StreamWriter( new PackagePartStream( styles_document.GetStream() ) ) )
        {
          style_document.Save( tw );
        }
      }
      #endregion

      var newXElement = new XElement( p.Xml );

      this.AddElementInXml( newXElement );

      int index = 0;
      if( this.Document._paragraphLookup.Keys.Count() > 0 )
      {
        index = this.Document._paragraphLookup.Last().Key;

        if( this.Document._paragraphLookup.Last().Value.Text.Length == 0 )
        {
          index++;
        }
        else
        {
          index += this.Document._paragraphLookup.Last().Value.Text.Length;
        }
      }

      var newParagraph = new Paragraph( Document, newXElement, index );
      this.Document._paragraphLookup.Add( index, newParagraph );
      this.SetParentContainer( newParagraph );
      return newParagraph;
    }

    public virtual Paragraph InsertParagraph( int index, string text, bool trackChanges, Formatting formatting )
    {
      var newParagraph = new Paragraph( this.Document, new XElement( Document.w + "p" ), index );
      newParagraph.InsertText( 0, text, trackChanges, formatting );

      var firstPar = HelperFunctions.GetFirstParagraphEffectedByInsert( Document, index );
      if( firstPar != null )
      {
        var splitIndex = index - firstPar._startIndex;
        if( splitIndex > 0 )
        {
          var splitParagraph = HelperFunctions.SplitParagraph( firstPar, splitIndex );
          firstPar.Xml.ReplaceWith( splitParagraph[ 0 ], newParagraph.Xml, splitParagraph[ 1 ] );
        }
        else
        {
          firstPar.Xml.ReplaceWith( newParagraph.Xml, firstPar.Xml );
        }
      }
      else
      {
        this.AddElementInXml( newParagraph );
      }

      this.SetParentContainer( newParagraph );
      return newParagraph;
    }

    public virtual Paragraph InsertParagraph( string text )
    {
      return InsertParagraph( text, false, new Formatting() );
    }

    public virtual Paragraph InsertParagraph( string text, bool trackChanges )
    {
      return InsertParagraph( text, trackChanges, new Formatting() );
    }

    public virtual Paragraph InsertParagraph( string text, bool trackChanges, Formatting formatting )
    {
      var newParagraph = new XElement
      (
          XName.Get( "p", Document.w.NamespaceName ), new XElement( XName.Get( "pPr", Document.w.NamespaceName ) ), HelperFunctions.FormatInput( text, formatting.Xml )
      );

      if( trackChanges )
      {
        newParagraph = HelperFunctions.CreateEdit( EditType.ins, DateTime.Now, newParagraph );
      }

      this.AddElementInXml( newParagraph );

      var newParagraphAdded = new Paragraph( this.Document, newParagraph, 0 );
      var cell = this as Cell;
      if( cell != null )
      {
        newParagraphAdded.PackagePart = cell.PackagePart;
      }
      else
      {
        var docx = this as Document;
        if( docx != null )
        {
          newParagraphAdded.PackagePart = this.Document.PackagePart;
        }
        else
        {
          var footer = this as Footer;
          if( footer != null )
          {
            newParagraphAdded.PackagePart = footer.PackagePart;
          }
          else
          {
            var header = this as Header;
            if( header != null )
            {
              newParagraphAdded.PackagePart = header.PackagePart;
            }
            else
            {
              newParagraphAdded.PackagePart = this.Document.PackagePart;
            }
          }
        }
      }

      this.SetParentContainer( newParagraphAdded );

      return newParagraphAdded;
    }

    /// <summary>
    /// Removes paragraph at specified position
    /// </summary>
    /// <param name="index">Index of paragraph to remove</param>
    /// <returns>True if paragraph removed</returns>
    public bool RemoveParagraphAt( int index )
    {
      var paragraphs = Xml.Descendants( Document.w + "p" ).ToList();
      if( index < paragraphs.Count )
      {
        paragraphs[ index ].Remove();
        return true;
      }

      return false;
    }

    /// <summary>
    /// Removes a paragraph
    /// </summary>
    /// <param name="paragraph">The paragraph to remove</param>
    /// <returns>True if paragraph removed</returns>
    public bool RemoveParagraph( Paragraph paragraph )
    {
      var paragraphs = Xml.Descendants( Document.w + "p" );
      var index = paragraphs.ToList().IndexOf( paragraph.Xml );

      if( index == -1 )
        return false;
      return this.RemoveParagraphAt( index );
    }

    public virtual Paragraph InsertEquation( string equation, Alignment align = Alignment.center )
    {
      Paragraph p = InsertParagraph();
      p.AppendEquation( equation, align );
      return p;
    }

    public virtual Paragraph InsertBookmark( String bookmarkName )
    {
      var p = InsertParagraph();
      p.AppendBookmark( bookmarkName );
      return p;
    }

    public virtual Table InsertTable( int rowCount, int columnCount )
    {
      var newTable = HelperFunctions.CreateTable( rowCount, columnCount );
      this.AddElementInXml( newTable );

      var table = new Table( this.Document, newTable, this.PackagePart );
      table.PackagePart = this.PackagePart;
      return table;
    }

    public virtual Table InsertTable( int index, int rowCount, int columnCount )
    {
      var newTable = HelperFunctions.CreateTable( rowCount, columnCount );

      var p = HelperFunctions.GetFirstParagraphEffectedByInsert( Document, index );
      if( p == null )
      {
        Xml.Elements().First().AddFirst( newTable );
      }
      else
      {
        var split = HelperFunctions.SplitParagraph( p, index - p._startIndex );
        p.Xml.ReplaceWith( split[ 0 ], newTable, split[ 1 ] );
      }

      var table = new Table( this.Document, newTable, this.PackagePart );
      table.PackagePart = this.PackagePart;
      return table;
    }

    public virtual Table InsertTable( Table t )
    {
      var newXElement = new XElement( t.Xml );
      this.AddElementInXml( newXElement );

      var newTable = new Table( this.Document, newXElement, this.PackagePart );
      newTable.Design = t.Design;
      newTable.PackagePart = this.PackagePart;
      return newTable;
    }

    public virtual Table InsertTable( int index, Table t )
    {
      var p = HelperFunctions.GetFirstParagraphEffectedByInsert( this.Document, index );

      var split = HelperFunctions.SplitParagraph( p, index - p._startIndex );
      var newXElement = new XElement( t.Xml );
      p.Xml.ReplaceWith( split[ 0 ], newXElement, split[ 1 ] );

      var newTable = new Table( this.Document, newXElement, this.PackagePart );
      newTable.Design = t.Design;
      newTable.PackagePart = this.PackagePart;
      return newTable;
    }

    public virtual Section InsertSection()
    {
      return this.InsertSection( false );
    }

    public virtual Section InsertSection( bool trackChanges )
    {
      return null;
    }

    public virtual Section InsertSectionPageBreak( bool trackChanges = false )
    {
      return null;
    }

    public virtual List InsertList( List list )
    {
      foreach( var item in list.Items )
      {
        this.AddElementInXml( item.Xml );
      }
      return list;
    }

    public virtual List InsertList( List list, double fontSize )
    {
      foreach( var item in list.Items )
      {
        item.FontSize( fontSize );
        this.AddElementInXml( item.Xml );
      }
      return list;
    }

    public virtual List InsertList( List list, Font fontFamily, double fontSize )
    {
      foreach( var item in list.Items )
      {
        item.Font( fontFamily );
        item.FontSize( fontSize );
        this.AddElementInXml( item.Xml );
      }
      return list;
    }

    public virtual List InsertList( int index, List list )
    {
      var p = HelperFunctions.GetFirstParagraphEffectedByInsert( Document, index );

      var split = HelperFunctions.SplitParagraph( p, index - p._startIndex );
      var elements = new List<XElement> { split[ 0 ] };
      elements.AddRange( list.Items.Select( i => new XElement( i.Xml ) ) );
      elements.Add( split[ 1 ] );
      p.Xml.ReplaceWith( elements.ToArray() );

      return list;
    }

    public virtual string[] ValidateBookmarks( params string[] bookmarkNames )
    {
      var result = new List<string>();

      foreach( var bookmarkName in bookmarkNames )
      {
        // Validate in container.
        if( this.Paragraphs.Any( p => p.ValidateBookmark( bookmarkName ) ) )
          return new string[ 0 ];

        result.Add( bookmarkName );
      }

      return result.ToArray();
    }

    public int RemoveTextInGivenFormat( Formatting formattingToMatch, MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch )
    {
      var count = 0;
      foreach( var element in Xml.Elements() )
        count += RecursiveRemoveText( element, formattingToMatch, formattingOptions );

      return count;
    }  

    #endregion

    #region Internal Methods

    protected internal virtual void AddElementInXml( object element )
    {
      this.Xml.Add( element );
    }

    internal List<Paragraph> GetParagraphs()
    {
      // Need some memory that can be updated by the recursive search.
      int index = 0;
      var paragraphs = new List<Paragraph>();

      var p = this.Xml.Descendants( XName.Get( "p", Document.w.NamespaceName ) );
      if( p != null )
      {
        foreach( XElement xElement in p )
        {
          // Do not include inner paragraphs contained in a Fallback.
          if( xElement.Ancestors().FirstOrDefault( x => x.Name.Equals( XName.Get( "Fallback", Document.mc.NamespaceName ) ) ) != null )
          {
            continue;
          }
          var paragraph = new Paragraph( this.Document, xElement, index );
          paragraph.PackagePart = this.PackagePart;
          paragraphs.Add( paragraph );
          index += HelperFunctions.GetText( xElement ).Length;
        }
      }

      return paragraphs;
    }

    internal void GetParagraphsRecursive( XElement xml, ref int index, ref List<Paragraph> paragraphs, bool isDeepSearch = false )
    {
      var keepSearching = true;
      if( xml.Name.LocalName == "p" )
      {
        paragraphs.Add( new Paragraph( Document, xml, index ) );

        index += HelperFunctions.GetText( xml ).Length;
        if( !isDeepSearch )
        {
          keepSearching = false;
        }
      }

      if( keepSearching && xml.HasElements )
      {
        foreach( XElement e in xml.Elements() )
        {
          this.GetParagraphsRecursive( e, ref index, ref paragraphs, isDeepSearch );
        }
      }
    }

    internal int RecursiveRemoveText( XElement element, Formatting formattingToMatch, MatchFormattingOptions formattingOptions )
    {
      var count = 0;
      foreach( var subElement in element.Elements() )
      {
        if( "rPr".Equals( subElement.Name.LocalName ) )
        {
          if( HelperFunctions.ContainsEveryChildOf( formattingToMatch.Xml, subElement, formattingOptions ) )
          {
            subElement.Parent.Remove();
            ++count;
          }
        }

        count += RecursiveRemoveText( subElement, formattingToMatch, formattingOptions );
      }

      return count;
    }

    #endregion

    #region Private Methods

    private void GetListItemType( Paragraph p )
    {
      var listItemType = HelperFunctions.GetListItemType( p, Document );
      if( listItemType != null )
      {
        p.ListItemType = GetListItemType( listItemType );
      }
    }

    private ContainerType GetParentFromXmlName( string xmlName )
    {
      switch( xmlName )
      {
        case "body":
          return ContainerType.Body;
        case "hdr":
          return ContainerType.Header;
        case "ftr":
          return ContainerType.Footer;
        case "p":
          return ContainerType.Paragraph;
        case "tbl":
          return ContainerType.Table;
        case "sectPr":
          return ContainerType.Section;
        case "tc":
          return ContainerType.Cell;
        case "txbxContent":
          return ContainerType.Shape;
        default:
          return ContainerType.None;
      }
    }

    private void SetParentContainer( Paragraph newParagraph )
    {
      var containerType = GetType();

      switch( containerType.Name )
      {
        case "Body":
          newParagraph.ParentContainer = ContainerType.Body;
          break;
        case "Table":
          newParagraph.ParentContainer = ContainerType.Table;
          break;
        case "TOC":
          newParagraph.ParentContainer = ContainerType.TOC;
          break;
        case "Section":
          newParagraph.ParentContainer = ContainerType.Section;
          break;
        case "Cell":
          newParagraph.ParentContainer = ContainerType.Cell;
          break;
        case "Header":
          newParagraph.ParentContainer = ContainerType.Header;
          break;
        case "Footer":
          newParagraph.ParentContainer = ContainerType.Footer;
          break;
        case "Paragraph":
          newParagraph.ParentContainer = ContainerType.Paragraph;
          break;
      }
    }

    private ListItemType GetListItemType( string listItemType )
    {
      switch( listItemType )
      {
        case "bullet":
          return ListItemType.Bulleted;
        default:
          return ListItemType.Numbered;
      }
    }

    private void InitParagraphs( List<Paragraph> paragraphs )
    {
      if( (this.Document.Sections == null) || (this.Document.Sections.Count == 0))
        return;

      foreach( var p in paragraphs )
      {
        var nextElement = p.Xml.ElementsAfterSelf().FirstOrDefault();
        if( ( nextElement == null ) && p.IsInSdt() )
        {
          nextElement = p.GetParentSdt().ElementsAfterSelf().FirstOrDefault();
        }
        var containsSectionBreak = p.GetOrCreate_pPr().Element( XName.Get( "sectPr", Document.w.NamespaceName ) );
        // Add FollowingTable to paragraph....only when paragraph is not the last one from a section.
        while( ( nextElement != null ) && ( nextElement.Name.Equals( Document.w + "tbl" ) ) && ( containsSectionBreak == null ) )
        {
          if( p.FollowingTables == null )
          {
            p.FollowingTables = new List<Table>();
          }
          p.FollowingTables.Add( new Table( this.Document, nextElement, this.PackagePart ) );
          nextElement = nextElement.ElementsAfterSelf().FirstOrDefault();
        }

        p.ParentContainer = this.GetParentFromXmlName( p.Xml.Ancestors().First().Name.LocalName );

        if( p.IsListItem )
        {
          this.GetListItemType( p );
        }
      }
    }

    #endregion

    #region Constructors

    internal Container( Document document, XElement xml )
        : base( document, xml )
    {

    }

    #endregion
  }
}
