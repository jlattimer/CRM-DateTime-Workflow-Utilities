using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class AddWeeksTests
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
        public void AddWeeks_Add_Plus_One()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "WeeksToAdd", 1 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2014, 7, 10, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddWeeks>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddWeeks_Add_Plus_Five()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "WeeksToAdd", 5 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2014, 8, 7, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddWeeks>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddWeeks_Add_Plus_Fifty()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "WeeksToAdd", 50 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2015, 6, 18, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddWeeks>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddWeeks_Add_Minus_One()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "WeeksToAdd", -1 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2014, 6, 26, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddWeeks>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddWeeks_Add_Minus_Five()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "WeeksToAdd", -5 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2014, 5, 29, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddWeeks>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddWeeks_Add_Minus_Fifty()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "WeeksToAdd", -50 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2013, 7, 18, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddWeeks>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }
    }
}