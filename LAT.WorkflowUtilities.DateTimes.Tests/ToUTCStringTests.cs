using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class ToUTCStringTests
    {
        #region Test Initialization and Cleanup
        // Use ClassInitialize to run code before running the first test in the class
        [ClassInitialize()]
        public static void ClassInitialize(TestContext testContext) { }

        // Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void ClassCleanup() { }

        // Use TestInitialize to run code before running each test 
        [TestInitialize()]
        public void TestMethodInitialize() { }

        // Use TestCleanup to run code after each test has run
        [TestCleanup()]
        public void TestMethodCleanup() { }
        #endregion

        [TestMethod]
        public void ToUTCString_Date_None()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToFormat", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "Culture", "" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "7/3/2014 1:48:00 PM";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToUTCString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        [TestMethod]
        public void ToUTCString_Date_enUS()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToFormat", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "Culture", "en-US" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "7/3/2014 1:48:00 PM";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToUTCString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        [TestMethod]
        public void ToUTCString_Date_jaJP()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToFormat", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "Culture", "ja-JP" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "2014/07/03 13:48:00";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToUTCString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        [TestMethod]
        [ExpectedException(typeof(System.Globalization.CultureNotFoundException))]
        public void ToUTCString_Date_Invalid()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToFormat", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "Culture", "x5x" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<ToUTCString>(workflowContext, inputs);
        }
    }
}