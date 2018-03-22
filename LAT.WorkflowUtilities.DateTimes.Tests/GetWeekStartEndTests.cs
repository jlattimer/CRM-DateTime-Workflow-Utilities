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
    public class GetWeekStartEndTests
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
        public void GetWeekStartEnd_Valid_Date1_User_Local()
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
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor(new DateTime(2014, 12, 31, 23, 0, 0));
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            DateTime expected1 = new DateTime(2014, 12, 28, 0, 0, 0);
            DateTime expected2 = new DateTime(2015, 1, 3, 23, 59, 59, 999);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetWeekStartEnd>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["WeekStartDate"]);
            Assert.AreEqual(expected2, result["WeekEndDate"]);
        }

        [TestMethod]
        public void GetWeekStartEnd_Valid_Date2_User_Local()
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
                { "DateToUse", new DateTime(2015, 5, 2, 23, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor(new DateTime(2015, 5, 3, 11, 0, 0));
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            DateTime expected1 = new DateTime(2015, 5, 3, 0, 0, 0);
            DateTime expected2 = new DateTime(2015, 5, 9, 23, 59, 59, 999);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetWeekStartEnd>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["WeekStartDate"]);
            Assert.AreEqual(expected2, result["WeekEndDate"]);
        }

        [TestMethod]
        public void GetWeekStartEnd__Valid_Date1_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2014, 12, 28, 0, 0, 0);
            DateTime expected2 = new DateTime(2015, 1, 3, 23, 59, 59, 999);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetWeekStartEnd>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["WeekStartDate"]);
            Assert.AreEqual(expected2, result["WeekEndDate"]);
        }

        [TestMethod]
        public void GetWeekStartEnd__Valid_Date2_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 5, 2, 23, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected1 = new DateTime(2015, 4, 26, 0, 0, 0);
            DateTime expected2 = new DateTime(2015, 5, 2, 23, 59, 59, 999);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetWeekStartEnd>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected1, result["WeekStartDate"]);
            Assert.AreEqual(expected2, result["WeekEndDate"]);
        }

        private class FakeLocalTimeFromUtcTimeRequestExecutor : IFakeMessageExecutor
        {
            private readonly DateTime _date;

            public FakeLocalTimeFromUtcTimeRequestExecutor(DateTime date)
            {
                _date = date;
            }

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
                ParameterCollection results = new ParameterCollection { { "LocalTime", _date } };
                localTime.Results = results;

                return localTime;
            }
        }
    }
}