using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class DateDiffMonthsTests
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
        public void DateDiffMonths_Minus_One_Month()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 6, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffMonths>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthsDifference"]);
        }

        [TestMethod]
        public void DateDiffMonths_Minus_Five_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 2, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffMonths>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthsDifference"]);
        }

        [TestMethod]
        public void DateDiffMonths_Minus_Twenty_Five_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2012, 6, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 25;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffMonths>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthsDifference"]);
        }

        [TestMethod]
        public void DateDiffMonths_Add_One_Month()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 8, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffMonths>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthsDifference"]);
        }

        [TestMethod]
        public void DateDiffMonths_Add_Five_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 12, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffMonths>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthsDifference"]);
        }

        [TestMethod]
        public void DateDiffMonths_Twenty_Five_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2016, 8, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 25;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffMonths>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthsDifference"]);
        }

        [TestMethod]
        public void DateDiffMonths_Same_Day()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffMonths>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthsDifference"]);
        }

        [TestMethod]
        public void DateDiffMonths_Minus_Eleven_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2013, 10, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 9;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffMonths>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthsDifference"]);
        }
    }
}