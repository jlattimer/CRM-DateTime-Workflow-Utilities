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
    public class DateDiffBusinessMinutes : WorkFlowActivityBase
    {
        public DateDiffBusinessMinutes() : base(typeof(DateDiffBusinessMinutes)) { }

        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [Input("Holiday/Closure Calendar")]
        [ReferenceTarget("calendar")]
        public InArgument<EntityReference> HolidayClosureCalendar { get; set; }

        [Output("Minutes Difference")]
        public OutArgument<int> MinutesDifference { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            var startDate = StartingDate.Get(context);
            startDate = new DateTime(startDate.Year, startDate.Month, startDate.Day, startDate.Hour, startDate.Minute, 0);
            var endDate = EndingDate.Get(context);
            endDate = new DateTime(endDate.Year, endDate.Month, endDate.Day, endDate.Hour, endDate.Minute, 0);

            if (startDate == endDate)
            {
                MinutesDifference.Set(context, 0);
                return;
            }

            Entity calendar = null;
            EntityReference holidayClosureCalendar = HolidayClosureCalendar.Get(context);
            if (holidayClosureCalendar != null)
                calendar = localContext.OrganizationService.Retrieve("calendar", holidayClosureCalendar.Id, new ColumnSet(true));

            var businessMinutes = 0;

            if (endDate > startDate)
            {
                while (startDate < endDate)
                {
                    if (startDate.IsBusinessMinute(calendar))
                    {
                        businessMinutes++;
                    }

                    startDate = startDate.AddMinutes(1);
                }
            }
            else
            {
                while (startDate > endDate)
                {
                    if (startDate.IsBusinessMinute(calendar))
                    {
                        businessMinutes--;
                    }

                    startDate = startDate.AddMinutes(-1);
                }
            }

            MinutesDifference.Set(context, businessMinutes);
        }
    }
}