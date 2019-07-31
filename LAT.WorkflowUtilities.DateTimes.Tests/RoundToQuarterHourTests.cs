using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class RoundToQuarterHourTests
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
        public void RoundUp_0_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 0, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 0, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, false);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundUp_12_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 12, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 15, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, false);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundUp_15_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 15, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 15, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, false);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundUp_20_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 20, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 30, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, false);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }


        [TestMethod]
        public void RoundUp_30_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 30, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 30, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, false);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundUp_35_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 35, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 45, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, false);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundUp_45_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 45, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 45, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, false);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundUp_55_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 55, 7, 0);

            var expected = new DateTime(2014, 7, 3, 9, 0, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, false);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundDown_0_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 0, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 0, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, true);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundDown_12_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 12, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 0, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, true);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundDown_15_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 15, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 15, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, true);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundDown_20_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 20, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 15, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, true);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }


        [TestMethod]
        public void RoundDown_30_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 30, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 30, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, true);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundDown_35_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 35, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 30, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, true);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundDown_45_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 45, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 45, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, true);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        [TestMethod]
        public void RoundDown_55_Minutes()
        {
            //Arrange
            var dateToRound = new DateTime(2014, 7, 3, 8, 55, 7, 0);

            var expected = new DateTime(2014, 7, 3, 8, 45, 0, 0);

            //Act
            var roundedDate = ExecuteWorkflow(dateToRound, true);

            //Assert
            Assert.AreEqual(expected, roundedDate);
        }

        private static DateTime ExecuteWorkflow(DateTime dateToRound, bool roundDown)
        {
            var workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToRound", dateToRound},
                { "RoundDown", roundDown }
            };

            var xrmFakedContext = new XrmFakedContext();

            var result = xrmFakedContext.ExecuteCodeActivity<RoundToQuarterHour>(workflowContext, inputs);

            return (DateTime)result["RoundedDate"];
        }
    }
}