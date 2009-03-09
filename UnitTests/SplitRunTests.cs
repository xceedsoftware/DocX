using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Novacode;
using System.Xml.Linq;

namespace UnitTests
{
    /// <summary>
    /// Summary description for SplitRunTests
    /// </summary>
    [TestClass]
    public class SplitRunTests
    {
        public SplitRunTests()
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
        public void TestSplitRun()
        {
            // The test text element to split
            Run r = new Run(0, new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "Hello world") }));

            #region Split at index 0
            /* 
             * Split r at index 0.
             * This will cause the left side of the split to be null and the right side to be equal to r.
            */
            XElement[] splitRun_indexZero = Run.SplitRun(r, 0);

            // Check if my expectations have been met
            Assert.IsNull(splitRun_indexZero[0]);
            Assert.AreEqual(r.Xml.ToString(), splitRun_indexZero[1].ToString());
            #endregion

            #region Split at index 1
            XElement[] splitRun_indexOne = Run.SplitRun(r, 1);

            // The result I expect to get from splitRun_indexOne
            XElement splitRun_indexOne_left = new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "H") });
            XElement splitRun_indexOne_right = new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "ello world") });

            // Check if my expectations have been met
            Assert.AreEqual(splitRun_indexOne_left.ToString(), splitRun_indexOne[0].ToString());
            Assert.AreEqual(splitRun_indexOne_right.ToString(), splitRun_indexOne[1].ToString());
            #endregion

            #region Split near the middle
            /* 
             * Split the text at index 11.
             * This will cause the left side of the split to end with a space and the right to start with a space.
            */
            XElement[] splitRun_nearMiddle = Run.SplitRun(r, 5);

            // The result I expect to get from splitRun_nearMiddle
            XElement splitRun_nearMiddle_left = new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "Hello") });
            XElement splitRun_nearMiddle_right = new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", new object[] { new XAttribute(XNamespace.Xml + "space", "preserve"), " world" }) });

            // Check if my expectations have been met
            Assert.AreEqual(splitRun_nearMiddle_left.ToString(), splitRun_nearMiddle[0].ToString());
            Assert.AreEqual(splitRun_nearMiddle_right.ToString(), splitRun_nearMiddle[1].ToString());
            #endregion

            #region Split at index Length - 1
            XElement[] splitRun_indexOneFromLength = Run.SplitRun(r, Paragraph.GetElementTextLength(r.Xml) - 1);

            // The result I expect to get from splitRun_indexOne
            XElement splitRun_indexOneFromLength_left = new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "Hello worl") });
            XElement splitRun_indexOneFromLength_right = new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "d") });

            // Check if my expectations have been met
            Assert.AreEqual(splitRun_indexOneFromLength_left.ToString(), splitRun_indexOneFromLength[0].ToString());
            Assert.AreEqual(splitRun_indexOneFromLength_right.ToString(), splitRun_indexOneFromLength[1].ToString());
            #endregion

            #region Split at index Length
            /* 
             * Split r at index Length.
             * This will cause the left side of the split to equal to r and the right side to be null.
            */
            XElement[] splitRun_indexLength = Run.SplitRun(r, Paragraph.GetElementTextLength(r.Xml));

            // Check if my expectations have been met
            Assert.AreEqual(r.Xml.ToString(), splitRun_indexLength[0].ToString());
            Assert.IsNull(splitRun_indexLength[1]); 
            #endregion
        }

        [TestMethod, ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void TestSplitRun_IndexLessThanTextStartIndex()
        {
            // The test text element to split
            Run r = new Run(0, new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "Hello world") }));
            
            /* 
             * Split r at a negative index.
             * This will cause an argument out of range exception to be thrown.
            */
            Run.SplitRun(r, r.StartIndex - 1);
        }

        [TestMethod, ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void TestSplitRun_IndexGreaterThanTextEndIndex()
        {
            // The test text element to split
            Run r = new Run(0, new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "t", "Hello world") }));
            
            /* 
             * Split r at a length + 1.
             * This will cause an argument out of range exception to be thrown.
            */

            Run.SplitRun(r, r.EndIndex + 1);   
        }

        [TestMethod]
        public void TestSplitRunOfLengthOne()
        {
            // The test text element to split
            Run r = new Run(0, new XElement(DocX.w + "r", new object[] { new XElement(DocX.w + "rPr", new object[] { new XElement(DocX.w + "b"), new XElement(DocX.w + "i"), new XElement(DocX.w + "color", new XAttribute(DocX.w + "val", "7030A0")) }), new XElement(DocX.w + "br") }));
            
            XElement[] splitRunOfLengthOne;

            #region Split before
            splitRunOfLengthOne = Run.SplitRun(r, r.StartIndex);
            // Check if my expectations have been met
            Assert.AreEqual(r.Xml.ToString(), splitRunOfLengthOne[0].ToString());
            Assert.IsNull(splitRunOfLengthOne[1]);
            #endregion

            #region Split after
            splitRunOfLengthOne = Run.SplitRun(r, r.EndIndex);
            // Check if my expectations have been met
            Assert.IsNull(splitRunOfLengthOne[0]);
            Assert.AreEqual(r.Xml.ToString(), splitRunOfLengthOne[1].ToString());
            #endregion
        }
    }
}
