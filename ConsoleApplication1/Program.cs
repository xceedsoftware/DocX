using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.IO.Packaging;
using System.Diagnostics;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {

            //// Testing constants
            //const string package_part_document = "/word/document.xml";
            //const string package_part_header_first = "/word/header3.xml";
            //const string package_part_header_odd = "/word/header2.xml";
            //const string package_part_header_even = "/word/header1.xml";
            //const string package_part_footer_first = "/word/footer3.xml";
            //const string package_part_footer_odd = "/word/footer2.xml";
            //const string package_part_footer_even = "/word/footer1.xml";

            //// Load Test-01.docx
            //using (DocX document = DocX.Load("../../data/Test-01.docx"))
            //{
            //    // Get the headers from the document.
            //    Headers headers = document.Headers; Debug.Assert(headers != null);
            //    Header header_first = headers.first; Debug.Assert(header_first != null);
            //    Header header_odd = headers.odd; Debug.Assert(header_odd != null);
            //    Header header_even = headers.even; Debug.Assert(header_even != null);

            //    // Get the footers from the document.
            //    Footers footers = document.Footers; Debug.Assert(footers != null);
            //    Footer footer_first = footers.first; Debug.Assert(footer_first != null);
            //    Footer footer_odd = footers.odd; Debug.Assert(footer_odd != null);
            //    Footer footer_even = footers.even; Debug.Assert(footer_even != null);

            //    // Its important that each Paragraph knows the PackagePart it belongs to.
            //    document.Paragraphs.ForEach(p => Debug.Assert(p.PackagePart.Uri.ToString() == package_part_document));
            //    header_first.Paragraphs.ForEach(p => Debug.Assert(p.PackagePart.Uri.ToString() == package_part_header_first));
            //    header_odd.Paragraphs.ForEach(p => Debug.Assert(p.PackagePart.Uri.ToString() == package_part_header_odd));
            //    header_even.Paragraphs.ForEach(p => Debug.Assert(p.PackagePart.Uri.ToString() == package_part_header_even));
            //    footer_first.Paragraphs.ForEach(p => Debug.Assert(p.PackagePart.Uri.ToString() == package_part_footer_first));
            //    footer_odd.Paragraphs.ForEach(p => Debug.Assert(p.PackagePart.Uri.ToString() == package_part_footer_odd));
            //    footer_even.Paragraphs.ForEach(p => Debug.Assert(p.PackagePart.Uri.ToString() == package_part_footer_even));
            //}
        }
    }
}
