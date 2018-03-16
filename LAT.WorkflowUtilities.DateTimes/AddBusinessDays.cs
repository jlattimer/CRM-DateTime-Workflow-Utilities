using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    using LAT.WorkflowUtilities.DateTimes.Common;

    public sealed class AddBusinessDays : CodeActivity
    {
        [RequiredArgument]
        [Input("Original Date")]
        public InArgument<DateTime> OriginalDate { get; set; }

        [RequiredArgument]
        [Input("Business Days To Add")]
        public InArgument<int> BusinessDaysToAdd { get; set; }

        [Input("Holiday/Closure Calendar")]
        [ReferenceTarget("calendar")]
        public InArgument<EntityReference> HolidayClosureCalendar { get; set; }

        [OutputAttribute("Updated Date")]
        public OutArgument<DateTime> UpdatedDate { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                DateTime originalDate = OriginalDate.Get(executionContext);
                int businessDaysToAdd = BusinessDaysToAdd.Get(executionContext);
                EntityReference holidaySchedule = HolidayClosureCalendar.Get(executionContext);

                DateTime tempDate = originalDate;

                if (businessDaysToAdd > 0)
                {
                    while (businessDaysToAdd > 0)
                    {
                        tempDate = tempDate.AddDays(1);

                        if (tempDate.IsBusinessDay(service, holidaySchedule))
                        {
                            // Only decrease the days to add if the day we've just added counts as a business day
                            businessDaysToAdd--;
                        }
                    }
                }
                else if (businessDaysToAdd < 0)
                {
                    while (businessDaysToAdd < 0)
                    {
                        tempDate = tempDate.AddDays(-1);

                        if (tempDate.IsBusinessDay(service, holidaySchedule))
                        {
                            // Only increase the days to add if the day we've just added counts as a business day
                            businessDaysToAdd++;
                        }
                    }
                }

                DateTime updatedDate = tempDate;

                UpdatedDate.Set(executionContext, updatedDate);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}