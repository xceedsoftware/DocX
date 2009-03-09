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
    /// Summary description for PreserveSpaceCheckTests
    /// </summary>
    [TestClass]
    public class PreserveSpaceTests
    {
        public PreserveSpaceTests()
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
        public void TestPreserveSpace_NoChangeRequired()
        {
            XElement t_origional, t_afterPreserveSpace;

            #region Doesn't start or end with a space
            // The test text element to check
            t_origional = new XElement(DocX.w + "t", "Hello I am  a string");
            t_afterPreserveSpace = t_origional;

            Text.PreserveSpace(t_afterPreserveSpace);
            Assert.AreEqual(t_origional.ToString(), t_afterPreserveSpace.ToString()); 
            #endregion

            #region Starts with a space, doesn't end with a space
            // The test text element to check
            t_origional = new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " Hello I am a string");
            t_afterPreserveSpace = t_origional;

            Text.PreserveSpace(t_afterPreserveSpace);
            Assert.AreEqual(t_origional.ToString(), t_afterPreserveSpace.ToString());
            #endregion

            #region Ends with a space, doesn't start with a space
            // The test text element to check
            t_origional = new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Hello I am a string ");
            t_afterPreserveSpace = t_origional;

            Text.PreserveSpace(t_afterPreserveSpace);
            Assert.AreEqual(t_origional.ToString(), t_afterPreserveSpace.ToString());
            #endregion

            #region Starts and ends with a space
            // The test text element to check
            t_origional = new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " Hello I am a string ");
            t_afterPreserveSpace = t_origional;

            Text.PreserveSpace(t_afterPreserveSpace);
            Assert.AreEqual(t_origional.ToString(), t_afterPreserveSpace.ToString());
            #endregion
        }

        [TestMethod]
        public void TestPreserveSpace_ChangeRequired()
        {
            XElement t_origional, t_afterPreserveSpace;

            #region Doesn't start or end with a space, but has a space preserve attribute
            // The test text element to check
            t_origional = new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Hello I am  a string");
            t_afterPreserveSpace = t_origional;

            Text.PreserveSpace(t_afterPreserveSpace);
            Assert.AreEqual((new XElement(DocX.w + "t", "Hello I am  a string")).ToString(), t_afterPreserveSpace.ToString());
            #endregion

            #region Start with a space, but has no space preserve attribute
            // The test text element to check
            t_origional = new XElement(DocX.w + "t", " Hello I am  a string");
            t_afterPreserveSpace = t_origional;

            Text.PreserveSpace(t_afterPreserveSpace);
            Assert.AreEqual((new XElement(DocX.w + "t",  new XAttribute(XNamespace.Xml + "space", "preserve"), " Hello I am  a string")).ToString(), t_afterPreserveSpace.ToString());
            #endregion

            #region Ends with a space, but has no space preserve attribute
            // The test text element to check
            t_origional = new XElement(DocX.w + "t", "Hello I am  a string ");
            t_afterPreserveSpace = t_origional;

            Text.PreserveSpace(t_afterPreserveSpace);
            Assert.AreEqual((new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Hello I am  a string ")).ToString(), t_afterPreserveSpace.ToString());
            #endregion

            #region Starts and ends with a space, but has no space preserve attribute
            // The test text element to check
            t_origional = new XElement(DocX.w + "t", " Hello I am  a string ");
            t_afterPreserveSpace = t_origional;

            Text.PreserveSpace(t_afterPreserveSpace);
            Assert.AreEqual((new XElement(DocX.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " Hello I am  a string ")).ToString(), t_afterPreserveSpace.ToString());
            #endregion
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void TestPreserveSpace_NotTOrDelTextElement()
        {
            XElement e = new XElement("NotTOrDelText");
            Text.PreserveSpace(e);
        }
    }
}
