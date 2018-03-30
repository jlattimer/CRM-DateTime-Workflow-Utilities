using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class DateDiffBusinessMinutesTests
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
        public void DateDiffBusinessMinutes_Minus_One_Day()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 3, 27, 0, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 3, 26, 0, 0, 0, 0)},
                { "HolidayClosureCalendar", null }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = -1440;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }

        [TestMethod]
        public void DateDiffBusinessMinutes_One_Hour()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 3, 27, 11, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 3, 27, 12, 0, 0, 0)},
                { "HolidayClosureCalendar", null }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = 60;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }

        [TestMethod]
        public void DateDiffBusinessMinutes_Minus_One_Hour()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 3, 27, 12, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 3, 27, 11, 0, 0, 0)},
                { "HolidayClosureCalendar", null }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = -60;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }

        [TestMethod]
        public void DateDiffBusinessMinutes_Same_DateTime()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 3, 27, 0, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 3, 27, 0, 0, 0, 0)},
                { "HolidayClosureCalendar", null }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }

        [TestMethod]
        public void DateDiffBusinessMinutes_Six_Days_With_Weekend_Between()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 3, 27, 0, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 4, 3, 0, 0, 0, 0)},
                { "HolidayClosureCalendar", null }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = 7200;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }

        [TestMethod]
        public void DateDiffBusinessMinutes_Minus_Six_Days_With_Weekend_Between()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 4, 3, 0, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 3, 27, 0, 0, 0, 0)},
                { "HolidayClosureCalendar", null }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = -7200;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }

        [TestMethod]
        public void DateDiffBusinessMinutes_Three_Days_Weekend_And_2_Full_Day_Holidays()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var calendarRules = new EntityCollection();

            var goodFridayCalendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            goodFridayCalendarRule.Attributes.Add("name", "Good Friday");
            goodFridayCalendarRule.Attributes.Add("starttime", new DateTime(2018, 3, 30, 0, 0, 0));
            goodFridayCalendarRule.Attributes.Add("duration", 1440);
            calendarRules.Entities.Add(goodFridayCalendarRule);

            var easterMondayCalendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            easterMondayCalendarRule.Attributes.Add("name", "Easter (Monday after)");
            easterMondayCalendarRule.Attributes.Add("starttime", new DateTime(2018, 4, 2, 0, 0, 0));
            easterMondayCalendarRule.Attributes.Add("duration", 1440);
            calendarRules.Entities.Add(easterMondayCalendarRule);

            var holidayCalendar = new Entity
            {
                Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 3, 27, 0, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 4, 3, 0, 0, 0, 0)},
                { "HolidayClosureCalendar", holidayCalendar.ToEntityReference() }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting, goodFridayCalendarRule, easterMondayCalendarRule, holidayCalendar });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = 4320;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }

        [TestMethod]
        public void DateDiffBusinessMinutes_Minus_Three_Days_Weekend_And_2_Full_Day_Holidays()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var calendarRules = new EntityCollection();

            var goodFridayCalendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            goodFridayCalendarRule.Attributes.Add("name", "Good Friday");
            goodFridayCalendarRule.Attributes.Add("starttime", new DateTime(2018, 3, 30, 0, 0, 0));
            goodFridayCalendarRule.Attributes.Add("duration", 1440);
            calendarRules.Entities.Add(goodFridayCalendarRule);

            var easterMondayCalendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            easterMondayCalendarRule.Attributes.Add("name", "Easter (Monday after)");
            easterMondayCalendarRule.Attributes.Add("starttime", new DateTime(2018, 4, 2, 0, 0, 0));
            easterMondayCalendarRule.Attributes.Add("duration", 1440);
            calendarRules.Entities.Add(easterMondayCalendarRule);

            var holidayCalendar = new Entity
            {
                Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 4, 3, 12, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 3, 29, 12, 0, 0, 0)},
                { "HolidayClosureCalendar", holidayCalendar.ToEntityReference() }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting, goodFridayCalendarRule, easterMondayCalendarRule, holidayCalendar });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = -1440;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }

        [TestMethod]
        public void DateDiffBusinessMinutes_Three_And_A_Half_Days_Half_Day_Holiday()
        {
            //Arrange
            var workflowContext = new XrmFakedWorkflowContext();

            var userSettings = new EntityCollection();
            var userSetting = new Entity("usersettings")
            {
                Id = Guid.NewGuid(),
                ["systemuserid"] = Guid.NewGuid(),
                ["timezonecode"] = 20
            };
            userSettings.Entities.Add(userSetting);

            var calendarRules = new EntityCollection();

            var halfDayCalendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            halfDayCalendarRule.Attributes.Add("name", "1/2 Day Closing");
            halfDayCalendarRule.Attributes.Add("starttime", new DateTime(2018, 3, 30, 12, 0, 0));
            halfDayCalendarRule.Attributes.Add("duration", 720);
            calendarRules.Entities.Add(halfDayCalendarRule);

            var holidayCalendar = new Entity
            {
                Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2018, 3, 27, 0, 0, 0, 0)},
                { "EndingDate", new DateTime(2018, 3, 31, 0, 0, 0, 0)},
                { "HolidayClosureCalendar", holidayCalendar.ToEntityReference() }
            };

            var xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting, halfDayCalendarRule, holidayCalendar });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));

            const int expected = 5040;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<DateDiffBusinessMinutes>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MinutesDifference"]);
        }
    }
}