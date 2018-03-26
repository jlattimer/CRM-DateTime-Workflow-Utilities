using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class RelativeTimeTests
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
        public void RelativeTime_One_Second_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 15, 1) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "one second ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Seconds_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 15, 10) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "10 seconds ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Minute_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 16, 2) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "a minute ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Minutes_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 30, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "15 minutes ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Hour_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 14, 14, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "an hour ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Hours_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 16, 17, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "3 hours ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Yesterday()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 2, 14, 14, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "yesterday";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Days_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 4, 14, 14, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "3 days ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_One_Month_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 6, 1, 14, 14, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "one month ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Months_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 7, 1, 14, 14, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "2 months ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_One_Year_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2016, 5, 1, 14, 14, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "one year ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Years_Ago()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2017, 5, 1, 14, 14, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "2 years ago";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }

        [TestMethod]
        public void RelativeTime_Future()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 30, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 15, 0) }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "in the future";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<RelativeTime>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["RelativeTimeString"]);
        }
    }
}