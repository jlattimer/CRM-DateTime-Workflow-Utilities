using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class AddBusinessDays : WorkFlowActivityBase
    {
        public AddBusinessDays() : base(typeof(AddBusinessDays)) { }

        [RequiredArgument]
        [Input("Original Date")]
        public InArgument<DateTime> OriginalDate { get; set; }

        [RequiredArgument]
        [Input("Business Days To Add")]
        public InArgument<int> BusinessDaysToAdd { get; set; }

        [Input("Holiday/Closure Calendar")]
        [ReferenceTarget("calendar")]
        public InArgument<EntityReference> HolidayClosureCalendar { get; set; }

        [Output("Updated Date")]
        public OutArgument<DateTime> UpdatedDate { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime originalDate = OriginalDate.Get(context);
            int businessDaysToAdd = BusinessDaysToAdd.Get(context);
            EntityReference holidaySchedule = HolidayClosureCalendar.Get(context);

            Entity calendar = null;
            if (holidaySchedule != null)
                calendar = localContext.OrganizationService.Retrieve("calendar", holidaySchedule.Id, new ColumnSet(true));

            DateTime tempDate = originalDate;

            if (businessDaysToAdd > 0)
            {
                while (businessDaysToAdd > 0)
                {
                    tempDate = tempDate.AddDays(1);

                    if (tempDate.IsBusinessDay(calendar))
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

                    if (tempDate.IsBusinessDay(calendar))
                    {
                        // Only increase the days to add if the day we've just added counts as a business day
                        businessDaysToAdd++;
                    }
                }
            }

            DateTime updatedDate = tempDate;

            UpdatedDate.Set(context, updatedDate);
        }
    }
}