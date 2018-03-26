using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class GetWeekStartEnd : WorkFlowActivityBase
    {
        public GetWeekStartEnd() : base(typeof(GetWeekStartEnd)) { }

        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Output("Week Start Date")]
        public OutArgument<DateTime> WeekStartDate { get; set; }

        [Output("Week End Date")]
        public OutArgument<DateTime> WeekEndDate { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime dateToUse = DateToUse.Get(context);
            bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(context);

            if (evaluateAsUserLocal)
            {
                int? timeZoneCode = GetLocalTime.RetrieveTimeZoneCode(localContext.OrganizationService);
                dateToUse = GetLocalTime.RetrieveLocalTimeFromUtcTime(dateToUse, timeZoneCode, localContext.OrganizationService);
            }

            int diff = dateToUse.DayOfWeek - DayOfWeek.Sunday;
            if (diff < 0)
                diff += 7;

            DateTime weekStartDate = dateToUse.AddDays(-1 * diff).Date;
            weekStartDate = new DateTime(weekStartDate.Year, weekStartDate.Month, weekStartDate.Day, 0, 0, 0);
            DateTime weekEndDate = weekStartDate.AddDays(6).AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);

            WeekStartDate.Set(context, weekStartDate);
            WeekEndDate.Set(context, weekEndDate);
        }
    }
}