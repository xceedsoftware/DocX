<img alt="DocX" src="https://download-codeplex.sec.s-msft.com/Download?ProjectName=docx&DownloadId=83756&Build=21031" />

[Codeplex](https://docx.codeplex.com/) | [GitHub](https://github.com/WordDocX/DocX) | [Blog](http://cathalscorner.blogspot.com/) |  [Donate](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=GHSZDFX6JHS2A&lc=GB&item_name=DocX%20library&currency_code=EUR&bn=PP%2dDonationsBF%3abtn_donateCC_LG_global%2egif%3aNonHosted)


<a href="https://www.nuget.org/packages/DocX/">
<img alt="NuGet Version" src="https://img.shields.io/nuget/v/DocX.svg" /> 
</a>
<a href="https://travis-ci.org/WordDocX/DocX">
<img alt="Travis CI Testing" src="https://travis-ci.org/WordDocX/DocX.svg?branch=master" />
</a>
<a href="https://ci.appveyor.com/project/PrzemyslawKlys/docx">
<img alt="AppVeyor Testing" src="https://ci.appveyor.com/api/projects/status/vxpnp8ivvvq2l39m?svg=true" />
</a>

***

## What is DocX?

DocX is a .NET library that allows developers to manipulate Word 2007/2010/2013 files, in an easy and intuitive manner. DocX is fast, lightweight and best of all it does not require Microsoft Word or Office to be installed.

DocX is available on both Codeplex and Github. We will try to update both.

## Install via Nuget
```
Install-Package DocX
```

## DocX Author

DocX was written by a PhD student **Cathal Coffey** studying at the National University of Ireland Maynooth. 
If you have found DocX useful and would like to buy Cathal lunch then you can do so via a paypal [donation](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=GHSZDFX6JHS2A&lc=GB&item_name=DocX%20library&currency_code=EUR&bn=PP%2dDonationsBF%3abtn_donateCC_LG_global%2egif%3aNonHosted).

To connect with Cathal on LinkedIn please follow: http://ie.linkedin.com/in/cathalcoffey

For Cathal's personal website follow: http://www.cathalcoffey.ie

## DocX Maintenance

Currently the development of DocX is mostly done by great support of community with project being maintained by Przemysław Kłys (MadBoy).

To connect with Przemek on LinkedIn please follow http://www.linkedin.com/in/pklys

For Przemek own little company website visit [Evotec](http://evotec.pl/)

## Cutting Edge

If you do not wish to wait for the next stable release of DocX.dll, you can build your own copy from the [latest source code](http://docx.codeplex.com/SourceControl/list/changesets#).

## What else do I need?

All that you need to install in order to use DocX is the [.NET framework 4.0](http://www.microsoft.com/downloads/en/details.aspx?FamilyID=9cfb2d51-5ff4-4491-b0e5-b386f32c0992&displaylang=en) and [Visual Studio 2010](http://www.microsoft.com/express/Downloads/) or later, both of which are free.

## What are the main features of DocX?

1. Insert, Remove or [Replace](http://cathalscorner.blogspot.com/2009/02/docx-net-library-for-manipulating-word.html) text in a document.  
All standard [text formatting](http://cathalscorner.blogspot.com/2009/08/docx-v1008-released.html) is available:
 1. Font {Family, Size, Color}, 
 2. Bold, 
 3. Italic, 
 4. Underline, 
 5. Strikethrough, 
 6. Script {Sub, Super}, 
 7. Highlight 
2. Here’s a [cool example](http://cathalscorner.blogspot.com/2010/12/replace-text-across-many-documents-in.html) of replacing text across many documents in Parallel 
3. Paragraph properties are exposed:
 1. Direction LeftToRight or RightToLeft,
 2. Indentation,
 3. Alignment
4. DocX also supports:
 1. [Pictures](http://cathalscorner.blogspot.com/2009/04/docx-version-1002-released.html), 
 2. [Hyperlinks](http://cathalscorner.blogspot.com/2010/06/docx-version-1009.html), 
 3. [Tables](http://cathalscorner.blogspot.com/2010/06/docx-and-tables.html), 
 4. [Headers & Footers](http://cathalscorner.blogspot.com/2010/06/docx-version-10010.html), 
 5. [Custom Properties](http://cathalscorner.blogspot.com/2009/02/docx-net-library-for-manipulating-word.html)

## Why would I use DocX?

DocX makes creating and manipulating documents a simple task. It does not use COM libraries nor does it require Microsoft Office to be installed. 

The following [blog post](http://cathalscorner.blogspot.com/2010/06/cathal-why-did-you-create-docx.html) compares the code used to create a HelloWorld document using:
 1. Office Interop libraries, 
 2. OOXML SDK, 
 3. DocX

## How can I learn more?

I have dedicated a blog to DocX. I regularly post new code examples [here](http://cathalscorner.blogspot.com/). The below videos are also excellent tutorials on how to use DocX.

[<img alt="Getting started" src="http://i3.codeplex.com/download?ProjectName=DocX&DownloadId=83768" />](http://docx.codeplex.com/Release/ProjectReleases.aspx?ReleaseId=32117#DownloadId=83636)
[<img alt="Paragraphs and text formatting" src="http://i3.codeplex.com/download?ProjectName=DocX&DownloadId=83995">](http://docx.codeplex.com/Release/ProjectReleases.aspx?ReleaseId=32117#DownloadId=83996)

## Advanced Examples

 1. Step by step guide to create an invoice for a company. http://cathalscorner.blogspot.com/2009/08/docx-v1007-released.html
 2. Replace text across many documents in Parallel. http://cathalscorner.blogspot.com/2010/12/replace-text-across-many-documents-in.html
 3. Programmatically manipulate an Image imbedded inside a document. http://cathalscorner.blogspot.com/2010/12/programmatically-manipulate-image.html
 4. Converting DocX into (.doc, .pdf, .html) http://cathalscorner.blogspot.com/2009/10/converting-docx-into-doc-pdf-html.html

Do you have an interesting or informative example that you would like to share? 
If you do, please email me.

## What do other people think?

<img alt="Testimonials" src="http://download.codeplex.com/download?ProjectName=DocX&DownloadId=192124">

## How can I send feedback!

If you have found DocX useful at work or in a personal project, I would love to hear about it. Equally if you have decided not to use DocX, please send me and email stating why this is so. I will use this feedback to improve DocX in future releases. 

My email address is coffey.cathal@gmail.com

## Other Projects

Cathal has another open source project [sql4csv](https://github.com/ccoffey/sql4csv/wiki). 
sql4csv is a python library that offers an SQL "like" interface for .csv files. 

## Our supporters

<a href="https://www.jetbrains.com/">
<img alt="ReSharper" src="https://evotec.xyz/resources/resharper_logos/logo_ReSharper.png" height = 100 />
</a>

***
