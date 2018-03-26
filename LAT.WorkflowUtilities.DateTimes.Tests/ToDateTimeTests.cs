using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class ToDateTimeTests
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
        public void ToDateTime_Valid_Date_Only1()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5/1/2015" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2015, 5, 1, 0, 0, 0);
            const bool expected2 = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }

        [TestMethod]
        public void ToDateTime_Valid_Date_Only2()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "05/01/2015" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2015, 5, 1, 0, 0, 0);
            const bool expected2 = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }

        [TestMethod]
        public void ToDateTime_Valid_Date_Only3()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5-1-2015" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2015, 5, 1, 0, 0, 0);
            const bool expected2 = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }

        [TestMethod]
        public void ToDateTime_Valid_Date_Time1()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5/1/2015 15:00:00" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2015, 5, 1, 15, 0, 0);
            const bool expected2 = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }

        [TestMethod]
        public void ToDateTime_Valid_Date_Time2()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5/1/2015 3:00 PM" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2015, 5, 1, 15, 0, 0);
            const bool expected2 = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }

        [TestMethod]
        public void ToDateTime_Valid_Date_Time3()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5-1-2015 15:00:00" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2015, 5, 1, 15, 0, 0);
            const bool expected2 = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }

        [TestMethod]
        public void ToDateTime_Valid_Date_Time4()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "Fri, 08 May 2015 21:45:58" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2015, 5, 8, 21, 45, 58);
            const bool expected2 = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }

        [TestMethod]
        public void ToDateTime_Valid_Date_Time5()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "2009-06-15T13:45:30" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2009, 6, 15, 13, 45, 30);
            const bool expected2 = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }

        [TestMethod]
        public void ToDateTime_InValid_Date()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "Hello World" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(1, 1, 1, 0, 0, 0);
            const bool expected2 = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<ToDateTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["ConvertedDate"]);
            Assert.AreEqual(expected2, result["IsValid"]);
        }
    }
}