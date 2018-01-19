## What is DocX?

DocX is a .NET library that allows developers to manipulate Word 2007/2010/2013 files, in an easy and intuitive manner. DocX is fast, lightweight and best of all it does not require Microsoft Word or Office to be installed.

**NOTE:** There is a new Master branch as of Oct. 3, 2017. Please read about the [Classic branch](../../wiki/Classic-Branch) if you were using this project before the change.

## DocX Author

DocX is the free, open source version of [Xceed Words for .NET](https://xceed.com/xceed-words-for-net). Originally written by Cathal Coffey, and maintained by Przemyslaw Klys, it is now maintained by Xceed. 

Currently, the only difference between DocX and Xceed Words for .NET, is that Xceed Words for .NET can convert a Word document to PDF, and has professional technical support included in the subscription.  

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
<td>$499.95</td>
</tr>
<tr>
<td>License</td>
<td>Ms-PL</td>
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

## What do other people think?

<img alt="Testimonials" src="http://download.codeplex.com/download?ProjectName=DocX&DownloadId=192124">

***

<a href="https://www.nuget.org/packages/DocX/">
<img alt="NuGet Version" src="https://img.shields.io/nuget/v/DocX.svg" /> 
</a>
