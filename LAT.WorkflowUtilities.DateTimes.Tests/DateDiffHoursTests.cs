using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class DateDiffHoursTests
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
        public void DateDiffHours_Minus_One_Hour()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 7, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Minus_Five_Hours()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 3, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Minus_Twenty_Five_Hours()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 7, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 25;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Minus_Two_Hunderd_Hours()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 6, 25, 0, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 200;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Add_One_Hour()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 9, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Add_Five_Hours()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 13, 48, 0, 00)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Add_Twenty_Five_Hours()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 4, 9, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 25;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Add_Two_Hunderd_Hours()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 11, 16, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 200;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Same_Day()
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
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }

        [TestMethod]
        public void DateDiffHours_Minus_Eleven_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2013, 10, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 6552;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffHours>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["HoursDifference"]);
        }
    }
}