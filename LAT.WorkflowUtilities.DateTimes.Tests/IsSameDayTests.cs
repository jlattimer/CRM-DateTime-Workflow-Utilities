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
    public class IsSameDayTests
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
        public void IsSameDay_Same_Day_User_Local()
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
                { "FirstDate", new DateTime(2015, 5, 1, 11, 36, 0) },
                { "SecondDate",  new DateTime(2015, 5, 1, 16, 49, 23) },
                { "EvaluateAsUserLocal", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor(new DateTime(2015, 5, 1, 5, 36, 0),
                new DateTime(2015, 5, 1, 11, 49, 23));
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const bool expected = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsSameDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["SameDay"]);
        }

        [TestMethod]
        public void IsSameDay_Not_Same_Day1_User_Local()
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
                { "FirstDate", new DateTime(2015, 5, 2, 11, 36, 0) },
                { "SecondDate",  new DateTime(2015, 5, 1, 16, 49, 23) },
                { "EvaluateAsUserLocal", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor(new DateTime(2015, 5, 2, 5, 36, 0),
                new DateTime(2015, 5, 1, 11, 49, 23));
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const bool expected = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsSameDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["SameDay"]);
        }

        [TestMethod]
        public void IsSameDay_Same_Day_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "FirstDate", new DateTime(2015, 5, 2, 11, 36, 0) },
                { "SecondDate",  new DateTime(2015, 5, 2, 16, 49, 23) },
                { "EvaluateAsUserLocal", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const bool expected = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsSameDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["SameDay"]);
        }

        [TestMethod]
        public void IsSameDay_Not_Same_Day_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "FirstDate", new DateTime(2015, 5, 1, 11, 36, 0) },
                { "SecondDate",  new DateTime(2015, 5, 2, 16, 49, 23) },
                { "EvaluateAsUserLocal", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const bool expected = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsSameDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["SameDay"]);
        }

        private class FakeLocalTimeFromUtcTimeRequestExecutor : IFakeMessageExecutor
        {
            private readonly DateTime _date1;
            private readonly DateTime _date2;
            private int _count = 0;

            public FakeLocalTimeFromUtcTimeRequestExecutor(DateTime date1, DateTime date2)
            {
                _date1 = date1;
                _date2 = date2;
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
                DateTime date = _date1;

                if (_count > 0)
                    date = _date2;

                OrganizationResponse localTime = new OrganizationResponse();
                ParameterCollection results = new ParameterCollection { { "LocalTime", date } };
                localTime.Results = results;

                _count++;

                return localTime;
            }
        }
    }
}