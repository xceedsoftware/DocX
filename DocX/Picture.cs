using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;

namespace Novacode
{
    /// <summary>
    /// Represents a Picture in this document, a Picture is a customized view of an Image.
    /// </summary>
    public class Picture: DocXElement
    {
        private const int EmusInPixel = 9525;

        internal Dictionary<PackagePart, PackageRelationship> picture_rels;
        
        internal Image img;
        private string id;
        private string name;
        private string descr;
        private int cx, cy;
        //private string fileName;
        private uint rotation;
        private bool hFlip, vFlip;
        private object pictureShape;
        private XElement xfrm;
        private XElement prstGeom;

        /// <summary>
        /// Remove this Picture from this document.
        /// </summary>
        public void Remove()
        {
            Xml.Remove();
        }

        /// <summary>
        /// Wraps an XElement as an Image
        /// </summary>
        /// <param name="i">The XElement i to wrap</param>
        internal Picture(DocX document, XElement i, Image img):base(document, i)
        {
            picture_rels = new Dictionary<PackagePart, PackageRelationship>();
            
            this.img = img;

            this.id =
            (
                from e in Xml.Descendants()
                where e.Name.LocalName.Equals("blip")
                select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value
            ).Single(); 

            this.name = 
            (
                from e in Xml.Descendants()
                let a = e.Attribute(XName.Get("name"))
                where (a != null)
                select a.Value
            ).First();
          
            this.descr =
            (
                from e in Xml.Descendants()
                let a = e.Attribute(XName.Get("descr"))
                where (a != null)
                select a.Value
            ).FirstOrDefault();

            this.cx = 
            (
                from e in Xml.Descendants()
                let a = e.Attribute(XName.Get("cx"))
                where (a != null)
                select int.Parse(a.Value)
            ).First();

            this.cy = 
            (
                from e in Xml.Descendants()
                let a = e.Attribute(XName.Get("cy"))
                where (a != null)
                select int.Parse(a.Value)
            ).First();

            this.xfrm =
            (
                from d in Xml.Descendants()
                where d.Name.LocalName.Equals("xfrm")
                select d
            ).Single();

            this.prstGeom =
            (
                from d in Xml.Descendants()
                where d.Name.LocalName.Equals("prstGeom")
                select d
            ).Single();

            this.rotation = xfrm.Attribute(XName.Get("rot")) == null ? 0 : uint.Parse(xfrm.Attribute(XName.Get("rot")).Value);
        }

        private void SetPictureShape(object shape)
        {
            this.pictureShape = shape;

            XAttribute prst = prstGeom.Attribute(XName.Get("prst"));
            if (prst == null)
                prstGeom.Add(new XAttribute(XName.Get("prst"), "rectangle"));

            prstGeom.Attribute(XName.Get("prst")).Value = shape.ToString();
        }

        /// <summary>
        /// Set the shape of this Picture to one in the BasicShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the BasicShapes enumeration.</param>
        public void SetPictureShape(BasicShapes shape)
        {
            SetPictureShape((object)shape);
        }

        /// <summary>
        /// Set the shape of this Picture to one in the RectangleShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the RectangleShapes enumeration.</param>
        public void SetPictureShape(RectangleShapes shape)
        {
            SetPictureShape((object)shape);
        }

        /// <summary>
        /// Set the shape of this Picture to one in the BlockArrowShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the BlockArrowShapes enumeration.</param>
        public void SetPictureShape(BlockArrowShapes shape)
        {
            SetPictureShape((object)shape);
        }

        /// <summary>
        /// Set the shape of this Picture to one in the EquationShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the EquationShapes enumeration.</param>
        public void SetPictureShape(EquationShapes shape)
        {
            SetPictureShape((object)shape);
        }

        /// <summary>
        /// Set the shape of this Picture to one in the FlowchartShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the FlowchartShapes enumeration.</param>
        public void SetPictureShape(FlowchartShapes shape)
        {
            SetPictureShape((object)shape);
        }

        /// <summary>
        /// Set the shape of this Picture to one in the StarAndBannerShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the StarAndBannerShapes enumeration.</param>
        public void SetPictureShape(StarAndBannerShapes shape)
        {
            SetPictureShape((object)shape);
        }

        /// <summary>
        /// Set the shape of this Picture to one in the CalloutShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the CalloutShapes enumeration.</param>
        public void SetPictureShape(CalloutShapes shape)
        {
            SetPictureShape((object)shape);
        }

        /// <summary>
        /// A unique id that identifies an Image embedded in this document.
        /// </summary>
        public string Id
        {
            get { return id; }
        }

        /// <summary>
        /// Flip this Picture Horizontally.
        /// </summary>
        public bool FlipHorizontal
        {
            get { return hFlip; }

            set
            {
                hFlip = value;

                XAttribute flipH = xfrm.Attribute(XName.Get("flipH"));
                if (flipH == null)
                    xfrm.Add(new XAttribute(XName.Get("flipH"), "0"));

                xfrm.Attribute(XName.Get("flipH")).Value = hFlip ? "1" : "0";
            }
        }

        /// <summary>
        /// Flip this Picture Vertically.
        /// </summary>
        public bool FlipVertical
        {
            get { return vFlip; }

            set
            {
                vFlip = value;

                XAttribute flipV = xfrm.Attribute(XName.Get("flipV"));
                if (flipV == null)
                    xfrm.Add(new XAttribute(XName.Get("flipV"), "0"));

                xfrm.Attribute(XName.Get("flipV")).Value = vFlip ? "1" : "0";
            }
        }

        /// <summary>
        /// The rotation in degrees of this image, actual value = value % 360
        /// </summary>
        public uint Rotation
        {
            get { return rotation / 60000; }

            set
            {
                rotation = (value % 360) * 60000;
                XElement xfrm =
                    (from d in Xml.Descendants()
                    where d.Name.LocalName.Equals("xfrm")
                    select d).Single();

                XAttribute rot = xfrm.Attribute(XName.Get("rot"));
                if(rot == null)
                    xfrm.Add(new XAttribute(XName.Get("rot"), 0));

                xfrm.Attribute(XName.Get("rot")).Value = rotation.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the name of this Image.
        /// </summary>
        public string Name 
        { 
            get { return name; } 
            
            set 
            { 
                name = value;

                foreach (XAttribute a in Xml.Descendants().Attributes(XName.Get("name")))
                    a.Value = name;
            } 
        }

        /// <summary>
        /// Gets or sets the description for this Image.
        /// </summary>
        public string Description 
        { 
            get { return descr; } 
            
            set 
            { 
                descr = value;

                foreach (XAttribute a in Xml.Descendants().Attributes(XName.Get("descr")))
                    a.Value = descr;
            } 
        }

        ///<summary>
        /// Returns the name of the image file for the picture.
        ///</summary>
        public string FileName
        {
          get
          {
            return img.FileName;
          }
        }

        /// <summary>
        /// Get or sets the Width of this Image.
        /// </summary>
        public int Width 
        { 
            get { return cx / EmusInPixel; }
            
            set 
            { 
                cx = value;

                foreach (XAttribute a in Xml.Descendants().Attributes(XName.Get("cx")))
                    a.Value = (cx * EmusInPixel).ToString();
            } 
        }

        /// <summary>
        /// Get or sets the height of this Image.
        /// </summary>
        public int Height 
        { 
            get { return cy / EmusInPixel; }
            
            set 
            { 
                cy = value;

                foreach (XAttribute a in Xml.Descendants().Attributes(XName.Get("cy")))
                    a.Value = (cy * EmusInPixel).ToString();
            } 
        }

        //public void Delete()
        //{
        //    // Remove xml
        //    i.Remove();
   
        //    // Rebuild the image collection for this paragraph
        //    // Requires that every Image have a link to its paragraph

        //}
    }
}
