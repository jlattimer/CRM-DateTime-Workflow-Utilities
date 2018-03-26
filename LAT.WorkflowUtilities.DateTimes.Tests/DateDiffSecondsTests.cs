using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class DateDiffSecondsTests
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
        public void DateDiffSeconds_Minus_One_Second()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 10, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 8, 48, 11, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffSeconds>(workflowContext, inputs);

            //Assertt
            Assert.AreEqual(expected, result["SecondsDifference"]);
        }

        [TestMethod]
        public void DateDiffSeconds_Minus_One_Hour()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 7, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 3600;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffSeconds>(workflowContext, inputs);

            //Assertt
            Assert.AreEqual(expected, result["SecondsDifference"]);
        }

        [TestMethod]
        public void DateDiffSeconds_Minus_Five_Hours()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 3, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 18000;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffSeconds>(workflowContext, inputs);

            //Assertt
            Assert.AreEqual(expected, result["SecondsDifference"]);
        }

        [TestMethod]
        public void DateDiffSeconds_Minus_Twenty_Five_Hours()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 7, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 90000;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffSeconds>(workflowContext, inputs);

            //Assertt
            Assert.AreEqual(expected, result["SecondsDifference"]);
        }

        [TestMethod]
        public void DateDiffSeconds_Add_One_Minute()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 8, 49, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 60;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffSeconds>(workflowContext, inputs);

            //Assertt
            Assert.AreEqual(expected, result["SecondsDifference"]);
        }

        [TestMethod]
        public void DateDiffSeconds_Add_Thirteen_Minute_With_Seconds()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 2, 21, 5, 47, 44, 0)},
                { "EndingDate", new DateTime(2014, 2, 21, 6, 00, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 736;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffSeconds>(workflowContext, inputs);

            //Assertt
            Assert.AreEqual(expected, result["SecondsDifference"]);
        }

        [TestMethod]
        public void DateDiffSeconds_Same_Day()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 10, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 8, 48, 10, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffSeconds>(workflowContext, inputs);

            //Assertt
            Assert.AreEqual(expected, result["SecondsDifference"]);
        }
    }
}