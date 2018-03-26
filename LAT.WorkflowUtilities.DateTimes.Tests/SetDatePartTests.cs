using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class SetDatePartTests
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
        public void SetDatePart_Valid_Hour()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "HR" },
                { "NewValue", 19 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2018, 3, 12, 19, 30, 20);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void SetDatePart_Valid_Minute()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "MIN" },
                { "NewValue", 19 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2018, 3, 12, 17, 19, 20);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void SetDatePart_Valid_Second()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "SEC" },
                { "NewValue", 19 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2018, 3, 12, 17, 30, 19);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void SetDatePart_Valid_Month()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "MON" },
                { "NewValue", 4 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2018, 4, 12, 17, 30, 20);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void SetDatePart_Valid_Day()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "DAY" },
                { "NewValue", 4 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2018, 3, 4, 17, 30, 20);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void SetDatePart_Valid_Year()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "YR" },
                { "NewValue", 2017 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2017, 3, 12, 17, 30, 20);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Invalid date part - must be HR/MIN/SEC MON/DAY/YR")]
        public void SetDatePart_Invalid_DatePart()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "TT" },
                { "NewValue", 2017 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Invalid hour")]
        public void SetDatePart_Invalid_Hour()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "HR" },
                { "NewValue", 2017 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Invalid minute")]
        public void SetDatePart_Invalid_Minute()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "MIN" },
                { "NewValue", 2017 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Invalid minute")]
        public void SetDatePart_Invalid_Second()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "SEC" },
                { "NewValue", -8 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Invalid month")]
        public void SetDatePart_Invalid_Month()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "MON" },
                { "NewValue", -8 }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Invalid day")]
        public void SetDatePart_Invalid_Day()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "DAY" },
                { "NewValue", 45}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Invalid year")]
        public void SetDatePart_Invalid_Year()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 3, 12, 17, 30, 20) },
                { "DatePart", "YR" },
                { "NewValue", 45}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidPluginExecutionException), "Update resulted in invalid date")]
        public void SetDatePart_Invalid_Date()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUpdate", new DateTime(2018, 2, 12, 17, 30, 20) },
                { "DatePart", "DAY" },
                { "NewValue", 31}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            //Act
            xrmFakedContext.ExecuteCodeActivity<SetDatePart>(workflowContext, inputs);
        }
    }
}