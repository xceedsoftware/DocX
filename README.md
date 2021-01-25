## Changes in this fork

This fork/branch includes several changes that were introduced in support of a project to generate "Register Format" reports from GEDCOM files. This use-case anticipates that the output document is to be processed programatically just once, when it is created, and will be subsequently edited manually.    

The changes include:

* Character Styles: if the format object passed when appending text to a paragraph includes a style name (id) it is applied to the text as a character style (in the base version, it is simply ignored)
* Indexing: supports inserting index entries (XE fields) into paragraphs and indexes (INDEX fields). Features supported include "type" (used to separate e.g. name index from place index)
* Footnotes/Endnotes: supports inserting footnotes and/or endnotes into paragraphs.  Notes can include hyperlinks and can include references to other notes (of the same flavor).  NB: This is a "one-way" capability: existing Notes (including ones just added using this feature) cannot be modified and cannot be deleted (except, of course, by hand in an editor).

Simple examples of footnotes, endnotes, and indexes are provided.

Full, unedited text of the base project readme follows.

## What is DocX?

DocX is a .NET library that allows developers to manipulate Word 2007/2010/2013 files, in an easy and intuitive manner. DocX is fast, lightweight and best of all it does not require Microsoft Word or Office to be installed.

**NOTE:** There is a new Master branch as of Oct. 3, 2017. Please read about the [Classic branch](../../wiki/Classic-Branch) if you were using this project before the change.

DocX is the free, open source version of [Xceed Words for .NET](https://xceed.com/xceed-words-for-net). Originally written by Cathal Coffey, and maintained by Przemyslaw Klys, it is now maintained by Xceed. 
Starting at v1.5.0, this free and open source product is provided under the Xceed Community License agreement(for non-commercial use). 

Currently, the differences between DocX and Xceed Words for .NET, is that Xceed Words for .NET :
* can convert a Word document to PDF
* adds properties to wrap text around Pictures/Tables/Shapes
* adds Picture cropping
* adds Shapes (rectangles for now)
* adds TextBoxes or Shapes containing Text
* gets Shapes from Paragraphs
* gets Charts from Paragraphs and can modify their categories/values
* is at least two versions ahead of the DocX version
* has professional technical support included in the subscription
* is available on .NET Standard 2.0 for .NET Core 2.0 Applications
* can automatically update fields from a document
* Insert html/rtf text (with tags), or html/rtf document, to a Word document
* Clone lists or tables
* Add or modify checkboxes
* Set transparency in pictures
* Create formatted hyperlinks based on a referenced hyperlinks
* Joining 2 documents gives the opportunity to choose the headers/footers of doc1, doc2 or both of them in the resulting document.

## What else do I need?

All that you need to install in order to use DocX is the [.NET framework 4.0](http://www.microsoft.com/downloads/en/details.aspx?FamilyID=9cfb2d51-5ff4-4491-b0e5-b386f32c0992&displaylang=en) and [Visual Studio 2010](http://www.microsoft.com/express/Downloads/) or later, both of which are free.

## What are the main features of DocX?

<table>

<tr>
<td>Edition</td>
<td><b>DocX</b></td>
<td><a href="https://xceed.com/xceed-words-for-net"><b>Xceed Words for .NET</b></a></td>
</tr>
<tr>
<td>Price</td>
<td>Free</td>
<td>$499.00</td>
</tr>
<tr>
<td>License</td>
<td>Xceed Community License</td>
<td>Proprietary</td>
</tr>
<tr>
<td>Email support</td>
<td></td>
<td>YES</td>
</tr>

<tr>
<td>Create new Word documents</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Modify Word documents</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Create new PDF documents</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Convert Word to PDF</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Supports .DOCX from Word 2007 and up</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Modify multiple documents in parallel for better performance</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Apply a template to a Word document</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Join documents, recreate portions from one to another</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Supports document protection with or without password</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Set document margins and page size</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Set line spacing, indentation, text direction, text alignment</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Wrap text around pictures</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Pictures with cropping</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Manage fonts and font sizes</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Set text color, bold, underline, italic, strikethrough, highlighting</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Set page numbering</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Create sections</td>
<td>YES</td>
<td>YES</td>
</tr>
<tr>
<td>Update document fields (ex: a table of contents) by calling only one method</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Wrap text around tables</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Wrap text around shapes</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Create shapes (rectangles for now)</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Create textboxes or shapes containing text</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Get shapes from paragraphs</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Get charts from paragraphs and modify their categories/values</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Update document fields with 1 method call</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Insert html/rtf text (with tags), or html/rtf document, to a Word document</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Clone lists or tables</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Add or modify checkboxes</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Set transparency in pictures</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Create formatted hyperlinks based on a referenced hyperlinks</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Joining 2 documents gives the opportunity to choose which headers/footers to use</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Available on .net standard 2.0+ for .net core 2.0+ applications</td>
<td></td>
<td>YES</td>
</tr>
<tr>
<td>Get release ahead</td>
<td></td>
<td>YES</td>
</tr>
</table>

**Supported Word document elements**

* Add headers or footers which can be the same on all pages, or unique for the first page, or unique for odd/even pages. Can contain images, hyperlinks and more.
* Insert/Modify paragraphs.
* Insert/Modify numbered or bulleted lists.
* Insert/Modify images. Flip, rotate, copy, modify, resize.
* Insert/Modify tables. Insert/Remove rows, columns, change direction, column width, row height, borders, merge/delete cells.
* Insert/Modify formatted equations or formulas.
* Insert/Modify bookmarks.
* Insert/Modify hyperlinks.
* Insert/Modify horizontal lines.
* Insert/Modify charts (bar, line, pie, 3D chart). Set colors, titles, legend, etc.
* Find, remove or replace text. Supports case sensitivity and regular expressions.
* Insert/Modify core or custom properties, such as author, address, subject, title, etc.
* Insert a Table Of Contents. Set title, change formatting.

## Why would I use DocX?

DocX makes creating and manipulating documents a simple task. It does not use COM libraries nor does it require Microsoft Office to be installed. 

The following [blog post](http://cathalscorner.blogspot.com/2010/06/cathal-why-did-you-create-docx.html) from Cathal Coffey compares the code used to create a HelloWorld document using:
 1. Office Interop libraries, 
 2. OOXML SDK, 
 3. DocX

## Advanced Examples

 1. Step by step guide to create an invoice for a company. http://cathalscorner.blogspot.com/2009/08/docx-v1007-released.html
 2. Replace text across many documents in Parallel. http://cathalscorner.blogspot.com/2010/12/replace-text-across-many-documents-in.html
 3. Programmatically manipulate an Image imbedded inside a document. http://cathalscorner.blogspot.com/2010/12/programmatically-manipulate-image.html
 4. Converting DocX into (.doc, .pdf, .html) http://cathalscorner.blogspot.com/2009/10/converting-docx-into-doc-pdf-html.html

Do you have an interesting or informative example that you would like to share? 
If you do, please email me.

## License Information

DocX is provided under the Xceed Software, Inc. Community License.

[<img src="https://user-images.githubusercontent.com/29377763/69274195-d9382200-0ba7-11ea-9ab7-bfce3126f35a.png"/>](license.md)

More information can be found in the [License](license.md) page.

A commercial license can be purchased at [Xceed](https://xceed.com).

## Release history

* **September 22, 2020, released DocX v1.7.1 with [19 bug fixes and improvements](../../wiki/Release-Notes-v1.7.1).**
* August 17, 2020, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.7.1 with [28 bug fixes and improvements](../../wiki/Release-Notes-v1.2.0#Plus171).
* **June 29, 2020, released DocX v1.7.0 with [27 bug fixes and improvements](../../wiki/Release-Notes-v1.7.0).**
* **January 30, 2020, released DocX v1.6.0 with [24 bug fixes and improvements](../../wiki/Release-Notes-v1.6.0).**
* January 30, 2020, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.7.0 with [62 bug fixes and improvements](../../wiki/Release-Notes-v1.2.0#Plus170).
* **November 26, 2019, released DocX v1.5.0 with [19 bug fixes and improvements](../../wiki/Release-Notes-v1.5.0).**
* **October 4, 2019, released DocX v1.4.1 with [12 bug fixes and improvements](../../wiki/Release-Notes-v1.4.1).**
* **February 21, 2019, released DocX v1.3.0 with [12 bug fixes and improvements](../../wiki/Release-Notes-v1.3.0).**
* January 28, 2019, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.6.0 with [71 bug fixes and improvements](../../wiki/Release-Notes-v1.2.0#Plus160).
* **June 27, 2018, released DocX v1.2.0 with [13 bug fixes and improvements](../../wiki/Release-Notes-v1.2.0).**
* June 18, 2018, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.5.0 with [71 bug fixes and improvements](../../wiki/Release-Notes-v1.1.0#Plus150).
* April 12, 2018, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.4.1 with [22 bug fixes and improvements](../../wiki/Release-Notes-v1.1.0#Plus141).
* January 15, 2018, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.4.0 with [19 bug fixes and improvements](../../wiki/Release-Notes-v1.1.0#Plus140).
* September 12, 2017, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.3.0 with [13 bug fixes and improvements](../../wiki/Release-Notes-v1.1.0#Plus130).
* June 5, 2017, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.2.0 with [13 bug fixes and improvements](../../wiki/Release-Notes-v1.1.0#Plus120).
* **October 3, 2017, released DocX v1.1.0 with [6 bug fixes and improvements](../../wiki/Release-Notes-v1.1.0).**
* March 1, 2017, released [Xceed Words for .NET](https://xceed.com/xceed-words-for-net) v1.1.0 with 6 bug fixes and improvements.

***

<a href="https://www.nuget.org/packages/DocX/">
<img alt="NuGet Version" src="https://img.shields.io/nuget/v/DocX.svg" /> 
</a>
