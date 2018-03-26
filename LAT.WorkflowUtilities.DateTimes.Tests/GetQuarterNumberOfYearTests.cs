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
    public class GetQuarterNumberOfYearTests
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
        public void GetQuarterNumberOfYear_Valid_Quarter_User_Local()
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
            var fakeLocalTimeFromUtcTimeRequestExecutor = new FakeLocalTimeFromUtcTimeRequestExecutor();
            xrmFakedContext.AddFakeMessageExecutor<LocalTimeFromUtcTimeRequest>(fakeLocalTimeFromUtcTimeRequestExecutor);

            const int expected = 4;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetQuarterNumberOfYear>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["QuarterNumberOfYear"]);
        }

        [TestMethod]
        public void GetQuarterNumberOfYear_Valid_Quarter_Utc()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<GetQuarterNumberOfYear>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["QuarterNumberOfYear"]);
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