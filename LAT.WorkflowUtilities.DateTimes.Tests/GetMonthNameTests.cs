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
    public class GetMonthNameTests
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
        public void GetMonthName_Valid_Month_User_Local_enUS()
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
                { "EvaluateAsUserLocal", true },
                { "Culture", "en-US"}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { userSetting });
            xrmFakedContext.CallerId = new EntityReference("systemuser", userSetting.GetAttributeValue<Guid>("systemuserid"));
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor();
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const string expected = "December";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetMonthName>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthName"]);
        }

        [TestMethod]
        public void GetMonthName_Valid_Month_Utc_enUS()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false },
                { "Culture", "en-US"}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "January";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetMonthName>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthName"]);
        }

        [TestMethod]
        public void GetMonthName_Valid_Month_Utc_esES()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false },
                { "Culture", "es-ES"}
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const string expected = "enero";

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetMonthName>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["MonthName"]);
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
                ParameterCollection results = new ParameterCollection { { "LocalTime", new DateTime(2014, 12, 31, 23, 0, 0) } };
                localTime.Results = results;

                return localTime;
            }
        }
    }
}