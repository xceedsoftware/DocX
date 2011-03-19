using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.Drawing;

namespace ConsoleApplication3
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a document.
            using (DocX document = DocX.Create(@"Test.docx"))
            {
                // Add a Table to this document.
                Table table = document.AddTable(2, 3);

                // Add a Hyperlink into this document.
                Hyperlink h = document.AddHyperlink("Google", new Uri("http://www.google.com"));

                // Add an Image into this document.
                Novacode.Image i = document.AddImage(@"C:\Users\cathal\Desktop\Logo.png");

                // Create a Picture (Custom View) of this Image.
                Picture p = i.CreatePicture();
                p.Rotation = 10;

                // Specify some properties for this Table.
                table.Alignment = Alignment.center;
                table.Design = TableDesign.LightShadingAccent2;

                // Insert the Table into the document.
                Table t1 = document.InsertTable(table);
                
                // Add content to this Table.
                t1.Rows[0].Cells[0].Paragraphs.First().AppendHyperlink(h).Append(" is my favourite search engine.");
                t1.Rows[0].Cells[1].Paragraphs.First().Append("This text is bold.").Bold();
                t1.Rows[0].Cells[2].Paragraphs.First().Append("Underlined").UnderlineStyle(UnderlineStyle.singleLine);
                t1.Rows[1].Cells[0].Paragraphs.First().Append("Green").Color(Color.Green);
                t1.Rows[1].Cells[1].Paragraphs.First().Append("Right to Left").Direction = Direction.RightToLeft;
                t1.Rows[1].Cells[2].Paragraphs.First().AppendPicture(p);

                document.Save();
            }// Release this document from memory.
        }
    }
}
