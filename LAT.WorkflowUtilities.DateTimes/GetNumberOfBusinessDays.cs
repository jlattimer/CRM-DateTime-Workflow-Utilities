using System;
using System.Activities;
using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public class GetNumberOfBusinessDays : WorkFlowActivityBase
    {
        public GetNumberOfBusinessDays() : base(typeof(GetNumberOfBusinessDays)) { }

        [RequiredArgument]
        [Input("Start Date For Range")]
        public InArgument<DateTime> StartDate { get; set; }

        [RequiredArgument]
        [Input("End Date For Range")]
        public InArgument<DateTime> EndDate { get; set; }

        [Input("Holiday/Closure Calendar")]
        [ReferenceTarget("calendar")]
        public InArgument<EntityReference> HolidayClosureCalendar { get; set; }

        [Output("Number Of Business Days")]
        public OutArgument<int> NumberOfBusinessDays { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            var dateToCheck = StartDate.Get(context);
            dateToCheck = new DateTime(dateToCheck.Year, dateToCheck.Month, dateToCheck.Day, 0, 0, 0);
            var endDate = EndDate.Get(context);
            endDate = new DateTime(endDate.Year, endDate.Month, endDate.Day, 0, 0, 0);

            var businessDays = 0;
            while (dateToCheck <= endDate)
            {
                if (dateToCheck.IsBusinessDay(localContext.OrganizationService, HolidayClosureCalendar.Get(context)))
                {
                    businessDays++;
                }

                dateToCheck = dateToCheck.AddDays(1);
            }

            NumberOfBusinessDays.Set(context, businessDays);
        }
    }
}