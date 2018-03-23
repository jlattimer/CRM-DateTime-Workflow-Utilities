using FakeXrmEasy;
using FakeXrmEasy.FakeMessageExecutors;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class GetFormattedDateStringTests
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
        public void GetFormattedDateString_2_Digit_24_Hour_Time_Only_User_Local()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection userSettings = new EntityCollection();
            Entity userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 15, 20, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", true },
                { "Format", "HH:mm" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor();
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const string expected = "23:15";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetFormattedDateString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        [TestMethod]
        public void GetFormattedDateString_2_Digit_24_Hour_Time_Only_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 15, 20, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false },
                { "Format", "HH:mm" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "05:15";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetFormattedDateString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        [TestMethod]
        public void GetFormattedDateString_2_Digit_With_Slashes_Date_Only_User_Local()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection userSettings = new EntityCollection();
            Entity userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 15, 20, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", true },
                { "Format", "MM/dd/yyyy" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor();
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const string expected = "12/31/2014";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetFormattedDateString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        [TestMethod]
        public void GetFormattedDateString_2_Digit_With_Slashes_Date_Only_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 15, 20, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false },
                { "Format", "MM/dd/yyyy" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "01/01/2015";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetFormattedDateString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        [TestMethod]
        public void GetFormattedDateString_2_Digit_With_Slashes_Date_And_Time_With_Milliseconds_User_Local()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection userSettings = new EntityCollection();
            Entity userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 15, 20, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", true },
                { "Format", "MM/dd/yyyy HH:mm:ss" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor();
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const string expected = "12/31/2014 23:15:20";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetFormattedDateString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        [TestMethod]
        public void GetFormattedDateString_2_Digit_With_Slashes_Date_And_Time_With_Milliseconds_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 15, 20, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false },
                { "Format", "MM/dd/yyyy HH:mm:ss" }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "01/01/2015 05:15:20";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetFormattedDateString>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["FormattedDateString"]);
        }

        private class FakeLocalTimeFromUtcTimeRequestExecutor : IFakeMessageExecutor
        {
            public bool CanExecute(OrganizationRequest request)
            {
                return request is LocalTimeFromUtcTimeRequest;
            }

            public Type GetResponsibleRequestType()
            {
                return typeof(LocalTimeFromUtcTimeRequest);
            }

            public OrganizationResponse Execute(OrganizationRequest request, XrmFakedContext ctx)
            {
                OrganizationResponse localTime = new OrganizationResponse();
                ParameterCollection results = new ParameterCollection { { "LocalTime", new DateTime(2014, 12, 31, 23, 15, 20) } };
                localTime.Results = results;

                return localTime;
            }
        }
    }
}