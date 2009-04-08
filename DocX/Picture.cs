using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace Novacode
{
    public class Picture
    {
        private string id;
        private string name;
        private string descr;
        private int cx, cy;
        private uint rotation;
        private bool hFlip, vFlip;
        private object pictureShape;

        // The underlying XElement which this Image wraps
        internal XElement i;

        private XElement xfrm;
        private XElement prstGeom;

        public Picture(string id, string name, string descr)
        {
            OpenXmlPart part = DocX.mainDocumentPart.GetPartById(id);

            this.id = id;
            this.name = name;
            this.descr = descr;

            using (System.Drawing.Image img = System.Drawing.Image.FromStream(part.GetStream()))
            {
                this.cx = img.Width * 4156;
                this.cy = img.Height * 4156;
            }

            XElement e = new XElement(DocX.w + "drawing");

            i = XElement.Parse 
            (string.Format(@"
            <drawing xmlns = ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                    <wp:extent cx=""{0}"" cy=""{1}"" />
                    <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0"" />
                    <wp:docPr id=""1"" name=""{3}"" descr=""{4}"" />
                    <wp:cNvGraphicFramePr>
                        <a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1"" />
                    </wp:cNvGraphicFramePr>
                    <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                        <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                            <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                <pic:nvPicPr>
                                <pic:cNvPr id=""0"" name=""{3}"" />
                                    <pic:cNvPicPr />
                                </pic:nvPicPr>
                                <pic:blipFill>
                                    <a:blip r:embed=""{2}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>
                                    <a:stretch>
                                        <a:fillRect />
                                    </a:stretch>
                                </pic:blipFill>
                                <pic:spPr>
                                    <a:xfrm>
                                        <a:off x=""0"" y=""0"" />
                                        <a:ext cx=""{0}"" cy=""{1}"" />
                                    </a:xfrm>
                                    <a:prstGeom prst=""rect"">
                                        <a:avLst />
                                    </a:prstGeom>
                                </pic:spPr>
                            </pic:pic>
                        </a:graphicData>
                    </a:graphic>
                </wp:inline>
            </drawing>
            ", cx, cy, id, name, descr));

            this.xfrm =
            (
                from d in i.Descendants()
                where d.Name.LocalName.Equals("xfrm")
                select d
            ).Single();

            this.prstGeom =
            (
                from d in i.Descendants()
                where d.Name.LocalName.Equals("prstGeom")
                select d
            ).Single();

            this.rotation = xfrm.Attribute(XName.Get("rot")) == null ? 0 : uint.Parse(xfrm.Attribute(XName.Get("rot")).Value);
        }

        /// <summary>
        /// Wraps an XElement as an Image
        /// </summary>
        /// <param name="i">The XElement i to wrap</param>
        internal Picture(XElement i)
        {
            this.i = i;

            this.id =
            (
                from e in i.Descendants()
                where e.Name.LocalName.Equals("blip")
                select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value
            ).Single();

            this.name = 
            (
                from e in i.Descendants()
                let a = e.Attribute(XName.Get("name"))
                where (a != null)
                select a.Value
            ).First();

            this.descr =
            (
                from e in i.Descendants()
                let a = e.Attribute(XName.Get("descr"))
                where (a != null)
                select a.Value
            ).First();

            this.cx = 
            (
                from e in i.Descendants()
                let a = e.Attribute(XName.Get("cx"))
                where (a != null)
                select int.Parse(a.Value)
            ).First();

            this.cy = 
            (
                from e in i.Descendants()
                let a = e.Attribute(XName.Get("cy"))
                where (a != null)
                select int.Parse(a.Value)
            ).First();

            this.xfrm =
            (
                from d in i.Descendants()
                where d.Name.LocalName.Equals("xfrm")
                select d
            ).Single();

            this.prstGeom =
            (
                from d in i.Descendants()
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

        public void SetPictureShape(BasicShapes shape)
        {
            SetPictureShape((object)shape);
        }

        public void SetPictureShape(RectangleShapes shape)
        {
            SetPictureShape((object)shape);
        }

        public void SetPictureShape(BlockArrowShapes shape)
        {
            SetPictureShape((object)shape);
        }

        public void SetPictureShape(EquationShapes shape)
        {
            SetPictureShape((object)shape);
        }

        public void SetPictureShape(FlowchartShapes shape)
        {
            SetPictureShape((object)shape);
        }

        public void SetPictureShape(StarAndBannerShapes shape)
        {
            SetPictureShape((object)shape);
        }

        public void SetPictureShape(CalloutShapes shape)
        {
            SetPictureShape((object)shape);
        }

        public string Id
        {
            get { return id; }
        }

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
                    (from d in i.Descendants()
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

                foreach (XAttribute a in i.Descendants().Attributes(XName.Get("name")))
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

                foreach (XAttribute a in i.Descendants().Attributes(XName.Get("descr")))
                    a.Value = descr;
            } 
        }

        /// <summary>
        /// Get or sets the Width of this Image.
        /// </summary>
        public int Width 
        { 
            get { return cx / 4156; }
            
            set 
            { 
                cx = value;

                foreach (XAttribute a in i.Descendants().Attributes(XName.Get("cx")))
                    a.Value = (cx * 4156).ToString();
            } 
        }

        /// <summary>
        /// Get or sets the height of this Image.
        /// </summary>
        public int Height 
        { 
            get { return cy / 4156; }
            
            set 
            { 
                cy = value;

                foreach (XAttribute a in i.Descendants().Attributes(XName.Get("cy")))
                    a.Value = (cy * 4156).ToString();
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
