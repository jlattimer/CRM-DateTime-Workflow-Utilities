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
    public class IsBusinessDayTests
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
        public void IsBusinessDay_Week_Day_User_Local()
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
                { "DateToCheck", new DateTime(2015, 5, 13, 3, 48, 0, 0)},
                { "HolidayClosureCalendar", null },
                { "EvaluateAsUserLocal", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor(new DateTime(2015, 5, 12, 9, 48, 0, 0));
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const bool expected = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBusinessDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["ValidBusinessDay"]);
        }

        [TestMethod]
        public void IsBusinessDay_Week_Day_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToCheck", new DateTime(2015, 5, 13, 3, 48, 0, 0)},
                { "HolidayClosureCalendar", null },
                { "EvaluateAsUserLocal", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const bool expected = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBusinessDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["ValidBusinessDay"]);
        }

        [TestMethod]
        public void IsBusinessDay_Week_End_User_Local()
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
                { "DateToCheck", new DateTime(2015, 5, 16, 3, 48, 0, 0)},
                { "HolidayClosureCalendar", null },
                { "EvaluateAsUserLocal", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor(new DateTime(2015, 5, 15, 9, 48, 0, 0));
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const bool expected = true;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBusinessDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["ValidBusinessDay"]);
        }

        [TestMethod]
        public void IsBusinessDay_Week_End_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToCheck", new DateTime(2015, 5, 16, 8, 48, 0, 0)},
                { "HolidayClosureCalendar", null },
                { "EvaluateAsUserLocal", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const bool expected = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBusinessDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["ValidBusinessDay"]);
        }

        [TestMethod]
        public void IsBusinessDay_Week_Day_Holiday_User_Local()
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

            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 0, 0, 0));
            calendarRule.Attributes.Add("duration", 1440);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "DateToCheck", new DateTime(2014, 7, 4, 8, 48, 0, 0)},
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70")}},
                { "EvaluateAsUserLocal", true }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting, calendarRule, holidayCalendar });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor(new DateTime(2014, 7, 4, 2, 48, 0, 0));
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const bool expected = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBusinessDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["ValidBusinessDay"]);
        }

        [TestMethod]
        public void IsBusinessDay_Week_Day_Holiday_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 0, 0, 0));
            calendarRule.Attributes.Add("duration", 1440);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");

            var inputs = new Dictionary<string, object>
            {
                { "DateToCheck", new DateTime(2014, 7, 5, 2, 48, 0, 0)},
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8")}},
                { "EvaluateAsUserLocal", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });

            const bool expected = false;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<IsBusinessDay>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["ValidBusinessDay"]);
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