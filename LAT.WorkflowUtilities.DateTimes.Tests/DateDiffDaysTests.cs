using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class DateDiffDaysTests
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
        public void DateDiffDays_Minus_One_Day()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Minus_Five_Days()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 6, 28, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Minus_Twenty_Five_Days()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 6, 8, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 25;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Minus_Two_Hunderd_Days()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2013, 12, 15, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 200;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Add_One_Day()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 4, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Add_Five_Days()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 8, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Add_Twenty_Five_Days()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 28, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 25;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Add_Two_Hunderd_Days()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2015, 1, 19, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 200;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Same_Day()
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
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        public void DateDiffDays_Minus_Eleven_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2013, 10, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 273;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["DaysDifference"]);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void DateDiffDays_Null_Date()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", null}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<DateDiffDays>(workflowContext, inputs);
        }
    }
}