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
        internal Uri uri;
        internal Dictionary<PackagePart, PackageRelationship> hyperlink_rels;

        /// <summary>
        /// Remove a Hyperlink from this Paragraph only.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///    // Add a hyperlink to this document.
        ///    Hyperlink h = document.AddHyperlink("link", new Uri("http://www.google.com"));
        ///
        ///    // Add a Paragraph to this document and insert the hyperlink
        ///    Paragraph p1 = document.InsertParagraph();
        ///    p1.Append("This is a cool ").AppendHyperlink(h).Append(" .");
        ///
        ///    /* 
        ///     * Remove the hyperlink from this Paragraph only. 
        ///     * Note a reference to the hyperlink will still exist in the document and it can thus be reused.
        ///     */
        ///    p1.Hyperlinks[0].Remove();
        ///
        ///    // Add a new Paragraph to this document and reuse the hyperlink h.
        ///    Paragraph p2 = document.InsertParagraph();
        ///    p2.Append("This is the same cool ").AppendHyperlink(h).Append(" .");
        ///
        ///    document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void Remove()
        {
            Xml.Remove();
        }

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
                List<XElement> newRuns = HelperFunctions.FormatInput(value, rPr);
                Xml.Add(newRuns);
            } 
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
                // Return the Hyperlinks Uri.
                return uri;
            } 

            set
            {
                foreach (PackagePart p in hyperlink_rels.Keys)
                {
                    PackageRelationship r = hyperlink_rels[p];

                    // Get all of the information about this relationship.
                    TargetMode r_tm = r.TargetMode;
                    string r_rt = r.RelationshipType;
                    string r_id = r.Id;

                    // Delete the relationship
                    p.DeleteRelationship(r_id);
                    p.CreateRelationship(value, r_tm, r_rt, r_id);
                }

                uri = value;
            } 
        }

        internal Hyperlink(DocX document, XElement i): base(document, i)
        {
            hyperlink_rels = new Dictionary<PackagePart, PackageRelationship>();
        }
    }
}
