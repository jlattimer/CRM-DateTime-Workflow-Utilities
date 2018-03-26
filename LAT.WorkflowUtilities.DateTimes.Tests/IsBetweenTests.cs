using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class IsBetweenTests
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
        public void IsBetween_Valid()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 11, 0, 0) },
                { "DateToValidate", new DateTime(2015, 5, 8, 15, 23, 0) },
                { "EndingDate", new DateTime(2015, 6, 1, 11, 0, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const bool expected = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBetween>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Between"]);
        }

        [TestMethod]
        public void IsBetween_Invalid_Equal()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 11, 0, 0) },
                { "DateToValidate", new DateTime(2015, 5, 8, 15, 23, 0) },
                { "EndingDate", new DateTime(2015, 5, 8, 15, 23, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const bool expected = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBetween>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Between"]);
        }

        [TestMethod]
        public void IsBetween_Invalid()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 11, 0, 0) },
                { "DateToValidate", new DateTime(2015, 8, 8, 15, 23, 0) },
                { "EndingDate", new DateTime(2015, 6, 1, 11, 0, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const bool expected = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBetween>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Between"]);
        }

        [TestMethod]
        public void IsBetween_Invalid_Invalid_Dates()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 6, 1, 11, 0, 0) },
                { "DateToValidate", new DateTime(2015, 8, 8, 15, 23, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 11, 0, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const bool expected = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBetween>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["Between"]);
        }
    }
}