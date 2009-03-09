using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Xml.Linq;
using Novacode;

namespace UnitTests
{
    /// <summary>
    /// Summary description for SplitEditTests
    /// </summary>
    [TestClass]
    public class SplitEditTests
    {
        public SplitEditTests()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void TestSplitEdit()
        {
            // The test text element to split
            XElement run1 = new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", "Hello"));
            XElement run2 = new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " world"));
            XElement edit = new XElement(DocX.w + "ins", new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), run1, run2);
            Paragraph p = new Paragraph(edit);

            #region Split at index 0
            XElement[] splitEdit = p.SplitEdit(edit, 0, EditType.del);

            Assert.IsNull(splitEdit[0]);
            Assert.AreEqual(edit.ToString(), splitEdit[1].ToString()); 
            #endregion

            #region Split at index 1
            /* 
             * Split the text at index 5.
             * This will cause the left side of the split to end with a space and the right to start with a space.
            */
            XElement[] splitEdit_indexOne = p.SplitEdit(edit, 1, EditType.del);

            // The result I expect to get from splitRun_nearMiddle
            XElement splitEdit_indexOne_left = new XElement(DocX.w + "ins", new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", "H")));
            XElement splitEdit_indexOne_right = new XElement(DocX.w + "ins", new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), new XElement(DocX.w + "r",new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", "ello")), new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " world")));

            // Check if my expectations have been met
            Assert.AreEqual(splitEdit_indexOne_left.ToString(), splitEdit_indexOne[0].ToString());
            Assert.AreEqual(splitEdit_indexOne_right.ToString(), splitEdit_indexOne[1].ToString());
            #endregion

            #region Split near the middle
            /* 
             * Split the text at index 5.
             * This will cause the left side of the split to end with a space and the right to start with a space.
            */
            XElement[] splitEdit_nearMiddle = p.SplitEdit(edit, 5, EditType.del);

            // The result I expect to get from splitRun_nearMiddle
            XElement splitEdit_nearMiddle_left = new XElement(DocX.w + "ins", new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", "Hello")));
            XElement splitEdit_nearMiddle_right = new XElement(DocX.w + "ins", new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " world")));
                
            // Check if my expectations have been met
            Assert.AreEqual(splitEdit_nearMiddle_left.ToString(), splitEdit_nearMiddle[0].ToString());
            Assert.AreEqual(splitEdit_nearMiddle_right.ToString(), splitEdit_nearMiddle[1].ToString());
            #endregion

            #region Split at index Length - 1
            /* 
             * Split the text at index 5.
             * This will cause the left side of the split to end with a space and the right to start with a space.
            */
            XElement[] splitEdit_indexOneFromLength = p.SplitEdit(edit, Paragraph.GetElementTextLength(edit) - 1, EditType.del);

            // The result I expect to get from splitRun_nearMiddle
            XElement splitEdit_OneFromLength_left = new XElement(DocX.w + "ins", new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", "Hello")), new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " worl")));
            XElement splitEdit_OneFromLength_right = new XElement(DocX.w + "ins", new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", "d")));

            // Check if my expectations have been met
            Assert.AreEqual(splitEdit_OneFromLength_left.ToString(), splitEdit_indexOneFromLength[0].ToString());
            Assert.AreEqual(splitEdit_OneFromLength_right.ToString(), splitEdit_indexOneFromLength[1].ToString());
            #endregion

            #region Split at index Length
            XElement[] splitEdit_indexZero = p.SplitEdit(edit, Paragraph.GetElementTextLength(edit), EditType.del);

            Assert.AreEqual(edit.ToString(), splitEdit_indexZero[0].ToString());
            Assert.IsNull(splitEdit_indexZero[1]);
            #endregion
        }

        [TestMethod, ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void TestSplitEdit_IndexLessThanTextStartIndex()
        {
            // The test text element to split
            XElement run1 = new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", "Hello"));
            XElement run2 = new XElement(DocX.w + "r", new XElement(DocX.w + "rPr", new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0"))), new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " world"));
            XElement edit = new XElement(DocX.w + "ins", new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), run1, run2);
            Paragraph p = new Paragraph(edit);

            /* 
             * Split r at a negative index.
             * This will cause an argument out of range exception to be thrown.
            */
            p.SplitEdit(edit, -1, EditType.del);
        }

        [TestMethod, ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void TestSplitEdit_IndexGreaterThanTextEndIndex()
        {
            // The test text element to split
            XElement run1 = new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "Hello") });
            XElement run2 = new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", new object[] { new XAttribute(XNamespace.Xml + "space", "preserve"), " world" }) });
            XElement edit = new XElement(DocX.w + "ins", new object[] { new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), run1, run2 });
            Paragraph p = new Paragraph(edit);

            /* 
             * Split r at a negative index.
             * This will cause an argument out of range exception to be thrown.
            */
            p.SplitEdit(edit, Paragraph.GetElementTextLength(edit) + 1, EditType.del);
        }

        [TestMethod]
        public void TestSplitEditOfLengthOne()
        {
            XElement edit = new XElement(DocX.w + "ins", new object[] { new XAttribute(DocX.w + "id", "0"), new XAttribute(DocX.w + "author", "t-cathco"), new XAttribute(DocX.w + "date", "2009-02-17T21:09:00Z"), new XElement(DocX.w + "r", new XElement(DocX.w + "tab"))});
            Paragraph p = new Paragraph(edit);

            XElement[] splitEditOfLengthOne;

            #region Split before
            splitEditOfLengthOne = p.SplitEdit(edit, 0, EditType.del);
            
            // Check if my expectations have been met
            Assert.AreEqual(edit.ToString(), splitEditOfLengthOne[0].ToString());
            Assert.IsNull(splitEditOfLengthOne[1]);
            #endregion

            #region Split after
            splitEditOfLengthOne = p.SplitEdit(edit, Paragraph.GetElementTextLength(edit), EditType.del);
            // Check if my expectations have been met
            Assert.IsNull(splitEditOfLengthOne[0]);
            Assert.AreEqual(edit.ToString(), splitEditOfLengthOne[1].ToString());
            #endregion
        }
    }
}
