using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class DateDiffTests
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
        public void DateDiff_Plus_1_Day()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 8, 48, 0, 0)},
                { "ShowSeconds", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "1d.00h:00m:00s";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiff>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Difference"]);
        }

        [TestMethod]
        public void DateDiff_Minus_1_Day()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 2, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "ShowSeconds", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "-1d.00h:00m:00s";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiff>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Difference"]);
        }

        [TestMethod]
        public void DateDiff_Plus_1_Hour()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 2, 9, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 8, 48, 0, 0)},
                { "ShowSeconds", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "0d.01h:00m:00s";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiff>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Difference"]);
        }

        [TestMethod]
        public void DateDiff_Plus_10_Hours_5_Minutes_With_Seconds()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 2, 18, 53, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 8, 48, 0, 0)},
                { "ShowSeconds", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "0d.10h:05m:00s";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiff>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Difference"]);
        }

        [TestMethod]
        public void DateDiff_Plus_10_Hours_5_Minutes_Without_Seconds()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 2, 18, 53, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 8, 48, 0, 0)},
                { "ShowSeconds", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "0d.10h:05m";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiff>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Difference"]);
        }

        [TestMethod]
        public void DateDiff_Plus_10_Hours_5_Minutes_14_Seconds()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 2, 18, 53, 14, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 8, 48, 0, 0)},
                { "ShowSeconds", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "0d.10h:05m:14s";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiff>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Difference"]);
        }
    }
}