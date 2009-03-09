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
    /// This class tests the SplitText function of Novacode.Paragraph
    /// </summary>
    [TestClass]
    public class SplitTextTests
    {
        public SplitTextTests()
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
        public void TestSplitText()
        {
            // The test text element to split
            Text t = new Text(0, new XElement(DocX.w + "t", "Hello I am  a string"));
            
            #region Split at index 0
            /* 
             * Split t at index 0.
             * This will cause the left side of the split to be null and the right side to be equal to t.
            */
            XElement[] splitText_indexZero = Text.SplitText(t, 0);

            // Check if my expectations have been met
            Assert.IsNull(splitText_indexZero[0]);
            Assert.AreEqual(t.Xml.ToString(), splitText_indexZero[1].ToString());
            #endregion

            #region Split at index 1
            XElement[] splitText_indexOne = Text.SplitText(t, 1);

            // The result I expect to get from splitText1
            XElement splitText_indexOne_left = new XElement(DocX.w + "t", "H");
            XElement splitText_indexOne_right = new XElement(DocX.w + "t", "ello I am  a string");

            // Check if my expectations have been met
            Assert.AreEqual(splitText_indexOne_left.ToString(), splitText_indexOne[0].ToString());
            Assert.AreEqual(splitText_indexOne_right.ToString(), splitText_indexOne[1].ToString());
            #endregion

            #region Split near the middle causing starting and ending spaces
            /* 
             * Split the text at index 11.
             * This will cause the left side of the split to end with a space and the right to start with a space.
            */
            XElement[] splitText_nearMiddle = Text.SplitText(t, 11);

            // The result I expect to get from splitText1
            XElement splitText_nearMiddle_left = new XElement(DocX.w + "t", new object[] { new XAttribute(XNamespace.Xml + "space", "preserve"), "Hello I am " });
            XElement splitText_nearMiddle_right = new XElement(DocX.w + "t", new object[] { new XAttribute(XNamespace.Xml + "space", "preserve"), " a string" });

            // Check if my expectations have been met
            Assert.AreEqual(splitText_nearMiddle_left.ToString(), splitText_nearMiddle[0].ToString());
            Assert.AreEqual(splitText_nearMiddle_right.ToString(), splitText_nearMiddle[1].ToString()); 
            #endregion

            #region Split at text.Value.Length - 1
            XElement[] splitText_indexOneFromLength = Text.SplitText(t, t.Value.Length - 1);

            // The result I expect to get from splitText1
            XElement splitText_indexOneFromLength_left = new XElement(DocX.w + "t", "Hello I am  a strin");
            XElement splitText_indexOneFromLength_right = new XElement(DocX.w + "t", "g");

            // Check if my expectations have been met
            Assert.AreEqual(splitText_indexOneFromLength_left.ToString(), splitText_indexOneFromLength[0].ToString());
            Assert.AreEqual(splitText_indexOneFromLength_right.ToString(), splitText_indexOneFromLength[1].ToString());
            #endregion

            #region Split at index text.Value.Length
            /* 
             * Split the text at index text.Value.Length.
             * This will cause the left side of the split to be equal to text and the right side to be null.
            */
            XElement[] splitText_indexLength = Text.SplitText(t, t.Value.Length);

            // Check if my expectations have been met
            Assert.AreEqual(t.Xml.ToString(), splitText_indexLength[0].ToString());
            Assert.IsNull(splitText_indexLength[1]);
            #endregion
        }

        [TestMethod, ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void TestSplitText_IndexLessThanTextStartIndex()
        {
            // The test text element to split
            Text t = new Text(0, new XElement(DocX.w + "t", "Hello I am  a string"));
            
            /* 
             * Split t at a negative index.
             * This will cause an argument out of range exception to be thrown.
            */
            Text.SplitText(t, t.StartIndex - 1);
        }

        [TestMethod, ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void TestSplitText_IndexGreaterThanTextEndIndex()
        {
            // The test text element to split
            Text t = new Text(0, new XElement(DocX.w + "t", "Hello I am  a string"));
            
            /* 
             * Split t at an index greater than its text length.
             * This will cause an argument out of range exception to be thrown.
            */
            Text.SplitText(t, t.EndIndex + 1);
        }

        [TestMethod]
        public void TestSplitTextOfLengthOne()
        {
            // The test text element to split
            Text t = new Text(0, new XElement(DocX.w + "tab"));
            XElement[] splitTextOfLengthOne;

            #region Split before
            splitTextOfLengthOne = Text.SplitText(t, t.StartIndex);
            // Check if my expectations have been met
            Assert.AreEqual(t.Xml.ToString(), splitTextOfLengthOne[0].ToString());
            Assert.IsNull(splitTextOfLengthOne[1]); 
            #endregion

            #region Split after
            splitTextOfLengthOne = Text.SplitText(t, t.EndIndex);
            // Check if my expectations have been met
            Assert.IsNull(splitTextOfLengthOne[0]);
            Assert.AreEqual(t.Xml.ToString(), splitTextOfLengthOne[1].ToString()); 
            #endregion
        }
    }
}
