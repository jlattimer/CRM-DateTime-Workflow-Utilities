using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class AddBusinessDaysTests
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
        public void AddBusinessDays_Add_One_No_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", null }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2014, 7, 4, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_Three_With_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 3 },
                { "HolidayClosureCalendar", null }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2014, 7, 8, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_Minus_One_No_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", -1 },
                { "HolidayClosureCalendar", null }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2014, 7, 2, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddMinusFiveWeekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", -5 },
                { "HolidayClosureCalendar", null }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            DateTime expected = new DateTime(2014, 6, 26, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_One_Non_Full_Day_With_Holiday()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 11, 0, 0));
            calendarRule.Attributes.Add("duration", 30);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = Guid.NewGuid(),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", holidayCalendar.ToEntityReference() }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });


            DateTime expected = new DateTime(2014, 7, 4, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_One_Non_Full_Day_With_Holiday_Plus_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 11, 0, 0));
            calendarRule.Attributes.Add("duration", 30);
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
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 4 },
                { "HolidayClosureCalendar", holidayCalendar.ToEntityReference() }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });

            DateTime expected = new DateTime(2014, 7, 9, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_One_Holiday_Plus_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

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
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", holidayCalendar.ToEntityReference() }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });

            DateTime expected = new DateTime(2014, 7, 7, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_One_Non_Full_Day_With_Business_Closure()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "Some Half Day");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 8, 0, 0));
            calendarRule.Attributes.Add("duration", 240);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Business Closure Calendar");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8")}}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });

            DateTime expected = new DateTime(2014, 7, 4, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_One_With_Business_Closure_Plus_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

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
                Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Business Closure Calendar");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8")}}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });

            DateTime expected = new DateTime(2014, 7, 7, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_Business_Closure_3_Days_Plus_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July extended 2nd-4th");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 2, 0, 0, 0));
            calendarRule.Attributes.Add("duration", 4320);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Business Closure Calendar");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2014, 7, 1, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8")}}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });

            DateTime expected = new DateTime(2014, 7, 7, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_Business_Closure_2_Days_No_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July extended 3rd-4th");
            calendarRule.Attributes.Add("starttime", new DateTime(2018, 7, 3, 0, 0, 0));
            calendarRule.Attributes.Add("duration", 2880);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Business Closure Calendar");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2018, 7, 2, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8")}}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });

            DateTime expected = new DateTime(2018, 7, 5, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }

        [TestMethod]
        public void AddBusinessDays_Add_Business_Closure_1_Full_Day_1_Partial_No_Weekend()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July extended 4th to 5th at 1200");
            calendarRule.Attributes.Add("starttime", new DateTime(2018, 7, 4, 0, 0, 0));
            calendarRule.Attributes.Add("duration", 2160);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Business Closure Calendar");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            var inputs = new Dictionary<string, object>
            {
                { "OriginalDate", new DateTime(2018, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8")}}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { calendarRule, holidayCalendar });

            // NOTE: While this is during the closure, the current logic only evaluates for full days - there for the 5th is a 'working day' 
            DateTime expected = new DateTime(2018, 7, 5, 8, 48, 0, 0);

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<AddBusinessDays>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UpdatedDate"]);
        }
    }
}