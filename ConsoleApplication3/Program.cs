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
            using (DocX document = DocX.Load(@"C:\Users\cathal\Downloads\foo.docx"))
            {
                List<Picture> pictures = document.Pictures;

                List<Novacode.Table> imageTable = (from table in document.Tables
                                                   where table.Pictures.Count > 0
                                                   select table).ToList();

            }// Release this document from memory.
        }
    }
}
