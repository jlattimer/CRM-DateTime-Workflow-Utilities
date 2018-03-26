using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class DateDiffYearsTests
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
        public void DateDiffYears_Minus_One_Year()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2013, 7, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }

        [TestMethod]
        public void DateDiffYears_Minus_Five_Years()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2009, 7, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }

        [TestMethod]
        public void DateDiffYears_Add_One_Year()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2015, 7, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }

        [TestMethod]
        public void DateDiffYears_Add_Five_Years()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2019, 7, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 5;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }

        [TestMethod]
        public void DateDiffYears_Add_One_Month()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 8, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }

        [TestMethod]
        public void DateDiffYears_Add_Eleven_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2015, 4, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }

        [TestMethod]
        public void DateDiffYears_Minus_One_Month()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 6, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }

        [TestMethod]
        public void DateDiffYears_Same_Day()
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
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }

        [TestMethod]
        public void DateDiffYears_Minus_Eleven_Months()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2013, 10, 3, 8, 48, 0, 0)}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffYears>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["YearsDifference"]);
        }
    }
}