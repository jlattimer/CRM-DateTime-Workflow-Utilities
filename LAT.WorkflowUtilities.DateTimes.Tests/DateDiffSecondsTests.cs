using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Moq;
using System;
using System.Activities;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class DateDiffSecondsTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public DateDiffSecondsTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.DateTimes.DateDiffSeconds" + ", " + "LAT.WorkflowUtilities.DateTimes";
        }
        #endregion
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
        public void MinusOneSecond()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 10, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 8, 48, 11, 0)}
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["SecondsDifference"]);
        }

        [TestMethod]
        public void MinusOneHour()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 7, 48, 0, 0)}
            };

            //Expected value
            const int expected = 3600;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["SecondsDifference"]);
        }

        [TestMethod]
        public void MinusFiveHours()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 3, 48, 0, 0)}
            };

            //Expected value
            const int expected = 18000;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["SecondsDifference"]);
        }

        [TestMethod]
        public void MinusTwentyFiveHours()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 2, 7, 48, 0, 0)}
            };

            //Expected value
            const int expected = 90000;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["SecondsDifference"]);
        }

        [TestMethod]
        public void OneMinute()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 8, 49, 0, 0)}
            };

            //Expected value
            const int expected = 60;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["SecondsDifference"]);
        }

        [TestMethod]
        public void ThirteenMinuteWithSeconds()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 2, 21, 5, 47, 44, 0)},
                { "EndingDate", new DateTime(2014, 2, 21, 6, 00, 0, 0)}
            };

            //Expected value
            const int expected = 736;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["SecondsDifference"]);
        }

        [TestMethod]
        public void Same()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2014, 7, 3, 8, 48, 10, 0)},
                { "EndingDate", new DateTime(2014, 7, 3, 8, 48, 10, 0)}
            };

            //Expected value
            const int expected = 0;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["SecondsDifference"]);
        }

        /// <summary>
        /// Invokes the workflow.
        /// </summary>
        /// <param name="name">Namespace.Class, Assembly</param>
        /// <param name="target">The target entity</param>
        /// <param name="inputs">The workflow input parameters</param>
        /// <param name="configuredServiceMock">The function to configure the Organization Service</param>
        /// <returns>The workflow output parameters</returns>
        private static IDictionary<string, object> InvokeWorkflow(string name, ref Entity target, Dictionary<string, object> inputs,
            Func<Mock<IOrganizationService>, Mock<IOrganizationService>> configuredServiceMock)
        {
            var testClass = Activator.CreateInstance(Type.GetType(name)) as CodeActivity; ;

            var serviceMock = new Mock<IOrganizationService>();
            var factoryMock = new Mock<IOrganizationServiceFactory>();
            var tracingServiceMock = new Mock<ITracingService>();
            var workflowContextMock = new Mock<IWorkflowContext>();

            //Apply configured Organization Service Mock
            if (configuredServiceMock != null)
                serviceMock = configuredServiceMock(serviceMock);

            IOrganizationService service = serviceMock.Object;

            //Mock workflow Context
            var workflowUserId = Guid.NewGuid();
            var workflowCorrelationId = Guid.NewGuid();
            var workflowInitiatingUserId = Guid.NewGuid();

            //Workflow Context Mock
            workflowContextMock.Setup(t => t.InitiatingUserId).Returns(workflowInitiatingUserId);
            workflowContextMock.Setup(t => t.CorrelationId).Returns(workflowCorrelationId);
            workflowContextMock.Setup(t => t.UserId).Returns(workflowUserId);
            var workflowContext = workflowContextMock.Object;

            //Organization Service Factory Mock
            factoryMock.Setup(t => t.CreateOrganizationService(It.IsAny<Guid>())).Returns(service);
            var factory = factoryMock.Object;

            //Tracing Service - Content written appears in output
            tracingServiceMock.Setup(t => t.Trace(It.IsAny<string>(), It.IsAny<object[]>())).Callback<string, object[]>(MoqExtensions.WriteTrace);
            var tracingService = tracingServiceMock.Object;

            //Parameter Collection
            ParameterCollection inputParameters = new ParameterCollection { { "Target", target } };
            workflowContextMock.Setup(t => t.InputParameters).Returns(inputParameters);

            //Workflow Invoker
            var invoker = new WorkflowInvoker(testClass);
            invoker.Extensions.Add(() => tracingService);
            invoker.Extensions.Add(() => workflowContext);
            invoker.Extensions.Add(() => factory);

            return invoker.Invoke(inputs);
        }
    }
}
