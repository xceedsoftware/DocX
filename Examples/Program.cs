using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Novacode;
using System.Text.RegularExpressions;
using System.Drawing;

namespace Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            StringReplaceExample();

            CustomPropertiesExample();

            CreateDocumentOnTheFly();

            ImageExample1();

            //ImageExample2();
        }

        private static void StringReplaceExample()
        {
            File.Copy(@"..\..\Input\StringReplace.docx", @"..\..\Output\StringReplace.docx", true);

            // Load the document that you want to manipulate
            DocX document = DocX.Load(@"..\..\Output\StringReplace.docx");

            // Loop through the paragraphs in the document
            foreach (Paragraph p in document.Paragraphs)
            {
                /* 
                 * Replace each instance of the string pear with the string banana.
                 * Specifying true as the third argument informs DocX to track the
                 * changes made by this replace. The fourth argument tells DocX to
                 * ignore case when matching the string pear.
                 */

                p.Replace("pear", "banana", true, RegexOptions.IgnoreCase);
            }

            // File will be saved to ..\..\Output\
            document.Save();
        }

        private static void CustomPropertiesExample()
        {
            // A list which contains three new users
            List<User> newUsers = new List<User>
            {
                new User 
                { 
                    forname = "John", username = "John87", 
                    freeGift = "toaster", joined = DateTime.Now, 
                    HomeAddress = "21 Hillview, Naas, Co. Kildare", 
                    RecieveFurtherMail = true
                },

                new User 
                {
                    forname = "James", username = "KingJames",
                    freeGift = "kitchen knife", joined = DateTime.Now, 
                    HomeAddress = "37 Mill Lane, Maynooth, Co. Meath", 
                    RecieveFurtherMail = false
                },

                new User 
                {
                    forname = "Mary", username = "McNamara1",
                    freeGift = "microwave", joined = DateTime.Now,  
                    HomeAddress = "110 Cherry Orchard Drive, Navan, Co. Roscommon", RecieveFurtherMail= true
                }
            };

            // Foreach of the three new user create a welcome document based on template.docx
            foreach (User newUser in newUsers)
            {
                // Copy template.docx so that we can customize it for this user
                string filename = string.Format(@"..\..\Output\{0}.docx", newUser.username);

                File.Copy(@"..\..\Input\Template.docx", filename, true);

                /* 
                 * Load the document to be manipulated and set the custom properties to this
                 * users specific data
                */
                DocX doc = DocX.Load(filename);
                doc.SetCustomProperty("Forname", CustomPropertyType.Text, newUser.forname);
                doc.SetCustomProperty("Username", CustomPropertyType.Text, newUser.username);
                doc.SetCustomProperty("FreeGift", CustomPropertyType.Text, newUser.freeGift);
                doc.SetCustomProperty("HomeAddress", CustomPropertyType.Text, newUser.HomeAddress);
                doc.SetCustomProperty("PleaseWaitNDays", CustomPropertyType.NumberInteger, 4);
                doc.SetCustomProperty("GiftArrivalDate", CustomPropertyType.Date, newUser.joined.AddDays(4).ToUniversalTime());
                doc.SetCustomProperty("RecieveFurtherMail", CustomPropertyType.YesOrNo, newUser.RecieveFurtherMail);

                // File will be saved to ..\..\Output\
                doc.Save();
                doc.Dispose();
            }
        }

        private static void CreateDocumentOnTheFly()
        {
            // Create a new .docx file
            DocX d = DocX.Create(@"..\..\Output\Hello.docx");

            // Add a new paragraph to this document
            Paragraph one = d.AddParagraph();

            one.Alignment = Alignment.both;

            // Create a text formatting called f1
            Formatting f1 = new Formatting();
            f1.FontFamily = new FontFamily("Agency FB");
            f1.Size = 28;
            f1.Bold = true;
            f1.FontColor = Color.RoyalBlue;
            f1.UnderlineStyle = UnderlineStyle.doubleLine;
            f1.UnderlineColor = Color.Red;

            // Insert some new text, into the new paragraph, using the created formatting f1
            one.Insert(0, "I've got style!", f1, false);

            // Create a text formatting called f2
            Formatting f2 = new Formatting();
            f2.FontFamily = new FontFamily("Colonna MT");
            f2.Size = 36.5;
            f2.Italic = true;
            f2.FontColor = Color.SeaGreen;

            // Insert some new text, into the new paragraph, using the created formatting f2
            one.Insert(one.Value.Length, " I have a different style.", f2, false);

            // Save the document
            d.Save();
            d.Dispose();
        }

        private static void ImageExample1()
        {
            File.Copy(@"..\..\Input\Image.docx", @"..\..\Output\Image.docx", true);

            // Load a .docx file
            DocX document = DocX.Load(@"..\..\Output\Image.docx");

            // Add an image to the docx file
            Novacode.Image img = document.AddImage(@"..\..\Input\Donkey.jpg");

            // Decide which paragraph to add the image to
            Paragraph p = document.Paragraphs.Last();

            #region pic1
            // Create a picture, a picture is a customized view of an image
            Picture pic1 = new Picture(img.Id, "Donkey", "Taken on Omey island");

            // Set the pictures shape
            pic1.SetPictureShape(BasicShapes.cube);

            // Rotate the picture clockwise by 30 degrees
            pic1.Rotation = 30;

            // Insert the picture at the end of this paragraph
            p.InsertPicture(pic1, p.Value.Length);
            #endregion

            #region pic2
            // Create a picture, a picture is a customized view of an image
            Picture pic2 = new Picture(img.Id, "Donkey", "Taken on Omey island");

            // Set the pictures shape
            pic2.SetPictureShape(CalloutShapes.cloudCallout);

            // Flip the picture horizontal
            pic2.FlipHorizontal = true;

            // Insert the picture at the end of this paragraph
            p.InsertPicture(pic2, p.Value.Length);
            #endregion

            // Save the docx file
            document.Save();
            document.Dispose();
        }

        private static void ImageExample2()
        {
            // Load the document that you want to manipulate
            DocX document = DocX.Load(@"..\..\Output\Image.docx");

            foreach (Paragraph p in document.Paragraphs)
            {
                foreach (Picture pi in p.Pictures)
                {
                    pi.Rotation = 30;
                }
            }

            // File will be saved to ..\..\Output\
            document.Save();
            document.Dispose();
        }
    }

    // This class is used in the CustomPropertiesExample()
    class User
    {
        public string forname, username, freeGift, HomeAddress;
        public DateTime joined;
        public bool RecieveFurtherMail;

        public User()
        { }
    }
}
