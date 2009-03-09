using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Novacode;

namespace CustomPropertyTestApp
{
    // This class represents a user
    class User
    {
        public string forname, username, freeGift, HomeAddress;
        public DateTime joined;
        public bool RecieveFurtherMail;

        public User()
        { }
    }

    class Program
    {
        static void Main(string[] args)
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
                string filename = string.Format(@"{0}.docx", newUser.username);

                File.Copy(@"Template.docx", filename, true);

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

                // File will be saved to \CustomPropertyTestApp\bin\Debug
                doc.Save();
            }
        }
    }
}
