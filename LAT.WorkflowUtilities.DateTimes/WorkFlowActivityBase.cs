using System;
using System.Activities;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;

// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    /// <inheritdoc />
    /// <summary>
    /// Base class for all workflow activity classes.
    /// </summary> 
    public abstract class WorkFlowActivityBase : CodeActivity
    {
        /// <summary>
        /// Workflow context object. 
        /// </summary>
        protected class LocalWorkflowContext
        {
            internal IServiceProvider ServiceProvider { get; private set; }

            /// <summary>
            /// The Microsoft Dynamics CRM organization service.
            /// </summary>
            internal IOrganizationService OrganizationService { get; }

            /// <summary>
            /// IWorkflowContext contains information that describes the run-time environment in which the workflow executes, information related to the execution pipeline, and entity business information.
            /// </summary>
            internal IWorkflowContext WorkflowExecutionContext { get; }

            /// <summary>
            /// Provides logging run-time trace information for plug-ins. 
            /// </summary>
            internal ITracingService TracingService { get; }

            private LocalWorkflowContext() { }

            /// <summary>
            /// Helper object that stores the services available in this workflow activity.
            /// </summary>
            /// <param name="executionContext"></param>
            internal LocalWorkflowContext(CodeActivityContext executionContext)
            {
                if (executionContext == null)
                {
                    throw new ArgumentNullException(nameof(executionContext));
                }

                // Obtain the execution context service from the service provider.
                WorkflowExecutionContext = executionContext.GetExtension<IWorkflowContext>();

                // Obtain the tracing service from the service provider.
                TracingService = executionContext.GetExtension<ITracingService>();

                // Obtain the Organization Service factory service from the service provider
                IOrganizationServiceFactory factory = executionContext.GetExtension<IOrganizationServiceFactory>();

                // Use the factory to generate the Organization Service.
                OrganizationService = factory.CreateOrganizationService(WorkflowExecutionContext.UserId);
            }

            /// <summary>
            /// Writes a trace message to the CRM trace log.
            /// </summary>
            /// <param name="message">Message name to trace.</param>
            internal void Trace(string message)
            {
                if (string.IsNullOrWhiteSpace(message) || TracingService == null)
                {
                    return;
                }

                if (WorkflowExecutionContext == null)
                {
                    TracingService.Trace(message);
                }
                else
                {
                    TracingService.Trace(
                        "{0}, Correlation Id: {1}, Initiating User: {2}",
                        message,
                        WorkflowExecutionContext.CorrelationId,
                        WorkflowExecutionContext.InitiatingUserId);
                }
            }
        }

        /// <summary>
        /// Gets or sets the name of the child class.
        /// </summary>
        /// <value>The name of the child class.</value>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "WorkFlowActivityBase")]
        private string ChildClassName { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorkFlowActivityBase"/> class.
        /// </summary>
        /// <param name="childClassName">The <see cref=" cred="Type"/> of the derived class.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "WorkFlowActivityBase")]
        internal WorkFlowActivityBase(Type childClassName)
        {
            ChildClassName = childClassName.ToString();
        }

        /// <inheritdoc />
        /// <summary>
        /// Main entry point for he business logic that the workflow activity is to execute.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics CRM caches plug-in instances. 
        /// The plug-in's Execute method should be written to be stateless as the constructor 
        /// is not called for every invocation of the plug-in. Also, multiple system threads 
        /// could execute the plug-in at the same time. All per invocation state information 
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        protected override void Execute(CodeActivityContext context)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            if (context == null)
                throw new ArgumentNullException(nameof(context));

            // Construct the Local workflow context.
            LocalWorkflowContext localContext = new LocalWorkflowContext(context);

            localContext.Trace(string.Format(CultureInfo.InvariantCulture, "Entered {0}.Execute()", ChildClassName));

#if DEBUG
            TraceArguments(true, context, localContext);
#endif

            try
            {
                ExecuteCrmWorkFlowActivity(context, localContext);

#if DEBUG
                TraceArguments(false, context, localContext);
#endif
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                localContext.Trace(string.Format(CultureInfo.InvariantCulture, "Exception: {0}", e.ToString()));

                // Handle the exception.
                throw;
            }
            finally
            {
                sw.Stop();
                localContext.Trace(string.Format(CultureInfo.InvariantCulture, "Exiting {0}.Execute() Duration: {1}ms", ChildClassName, sw.ElapsedMilliseconds));
            }
        }

        protected virtual void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            // Do nothing. 
        }

        private void TraceArguments(bool input, CodeActivityContext context, LocalWorkflowContext localContext)
        {
            try
            {
                var properties = GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);

                properties.ToList().ForEach(p =>
                {
                    if (input)
                    {
                        if (!p.PropertyType.IsSubclassOf(typeof(InArgument)) && !p.PropertyType.IsSubclassOf(typeof(InOutArgument)))
                            return;
                    }
                    else
                    {
                        if (!p.PropertyType.IsSubclassOf(typeof(OutArgument)) && !p.PropertyType.IsSubclassOf(typeof(InOutArgument)))
                            return;
                    }

                    var propertyLabel = input
                        ? ((InputAttribute)p.GetCustomAttribute(typeof(InputAttribute))).Name
                        : ((OutputAttribute)p.GetCustomAttribute(typeof(OutputAttribute))).Name;

                    StringBuilder sb = new StringBuilder();
                    sb.Append($"{p.PropertyType.BaseType?.Name}({p.PropertyType.GenericTypeArguments[0].FullName}): {propertyLabel}: ");

                    var property = (Argument)p.GetValue(this);
                    var propertyValue = property.Get(context);
                    switch (propertyValue)
                    {
                        case null:
                            sb.Append("Null");
                            break;
                        case string _:
                        case decimal _:
                        case int _:
                        case bool _:
                            sb.Append(propertyValue.ToString());
                            break;
                        case DateTime _:
                            sb.Append(((DateTime)propertyValue).ToString("yyyy-MM-dd HH:mm:ss \"GMT\"Z"));
                            break;
                        case EntityReference _:
                            var er = (EntityReference)propertyValue;
                            sb.Append($"Id: {er.Id}, LogicalName: {er.LogicalName}");
                            break;
                        case OptionSetValue _:
                            sb.Append(((OptionSetValue)propertyValue).Value);
                            break;
                        case Money _:
                            sb.Append(((Money)propertyValue).Value.ToString(CultureInfo.CurrentCulture));
                            break;
                        default:
                            sb.Append($"Undefined Type - {p.GetType().FullName}");
                            break;
                    }

                    localContext.TracingService.Trace(sb.ToString());
                });
            }
            catch (Exception e)
            {
                localContext.TracingService.Trace($"Error tracing arguments: {e.Message}");
            }
        }
    }
}