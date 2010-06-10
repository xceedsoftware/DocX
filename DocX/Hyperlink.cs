using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO.Packaging;

namespace Novacode
{
    /// <summary>
    /// Represents a Hyperlink in a document.
    /// </summary>
    public class Hyperlink: DocXElement
    {
        /// <summary>
        /// Change the Text of a Hyperlink.
        /// </summary>
        /// <example>
        /// Change the Text of a Hyperlink.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///    // Get all of the hyperlinks in this document
        ///    List<Hyperlink> hyperlinks = document.Hyperlinks;
        ///    
        ///    // Change the first hyperlinks text and Uri
        ///    Hyperlink h0 = hyperlinks[0];
        ///    h0.Text = "DocX";
        ///    h0.Uri = new Uri("http://docx.codeplex.com");
        ///
        ///    // Save this document.
        ///    document.Save();
        /// }
        /// </code>
        /// </example>
        public string Text 
        { 
            get
            {
                // Create a string builder.
                StringBuilder sb = new StringBuilder();

                // Get all the runs in this Text.
                var runs = from r in Xml.Elements()
                           where r.Name.LocalName == "r"
                           select new Run(Document, r, 0);

                // Remove each run.
                foreach (Run r in runs)
                    sb.Append(r.Value);

                return sb.ToString();
            } 

            set
            {
                // Get all the runs in this Text.
                var runs = from r in Xml.Elements()
                           where r.Name.LocalName == "r"
                           select r;

                // Remove each run.
                for (int i = 0; i < runs.Count(); i++)
                    runs.Remove();

                XElement rPr = 
                new XElement
                (
                    DocX.w + "rPr",
                    new XElement
                    (
                        DocX.w + "rStyle",
                        new XAttribute(DocX.w + "val", "Hyperlink")
                    )
                );

                // Format and add the new text.
                List<XElement> newRuns = DocX.FormatInput(value, rPr);
                Xml.Add(newRuns);
            } 
        }

        // Return the Id of this Hyperlink.
        internal string GetId()
        {
            return Xml.Attribute(DocX.r + "id").Value;
        }

        /// <summary>
        /// Change the Uri of a Hyperlink.
        /// </summary>
        /// <example>
        /// Change the Uri of a Hyperlink.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///    // Get all of the hyperlinks in this document
        ///    List<Hyperlink> hyperlinks = document.Hyperlinks;
        ///    
        ///    // Change the first hyperlinks text and Uri
        ///    Hyperlink h0 = hyperlinks[0];
        ///    h0.Text = "DocX";
        ///    h0.Uri = new Uri("http://docx.codeplex.com");
        ///
        ///    // Save this document.
        ///    document.Save();
        /// }
        /// </code>
        /// </example>
        public Uri Uri 
        { 
            get
            {
                // Get the word\document.xml part
                PackagePart word_document = Document.package.GetPart(new Uri("/word/document.xml", UriKind.Relative));

                // Get the Hyperlink relation based on its Id.
                PackageRelationship r = word_document.GetRelationship(GetId());
                
                // Return the Hyperlinks Uri.
                return r.TargetUri;
            } 

            set
            {
                // Get the word\document.xml part
                PackagePart word_document = Document.package.GetPart(new Uri("/word/document.xml", UriKind.Relative));

                // Get the Hyperlink relation based on its Id.
                PackageRelationship r = word_document.GetRelationship(GetId());

                // Get all of the information about this relationship.
                TargetMode r_tm = r.TargetMode;
                string r_rt = r.RelationshipType;
                string r_id = r.Id;

                // Delete the relationship
                word_document.DeleteRelationship(r_id);
                word_document.CreateRelationship(value, r_tm, r_rt, r_id);     
            } 
        }

        internal Hyperlink(DocX document, XElement i): base(document, i)
        {

        }
    }
}
