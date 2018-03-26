using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class GetMonthStartEnd : WorkFlowActivityBase
    {
        public GetMonthStartEnd() : base(typeof(GetMonthStartEnd)) { }

        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Output("Month Start Date")]
        public OutArgument<DateTime> MonthStartDate { get; set; }

        [Output("Month End Date")]
        public OutArgument<DateTime> MonthEndDate { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime dateToUse = DateToUse.Get(context); ;
            bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(context);

            if (evaluateAsUserLocal)
            {
                int? timeZoneCode = GetLocalTime.RetrieveTimeZoneCode(localContext.OrganizationService);
                dateToUse = GetLocalTime.RetrieveLocalTimeFromUtcTime(dateToUse, timeZoneCode, localContext.OrganizationService);
            }

            DateTime monthStartDate = new DateTime(dateToUse.Year, dateToUse.Month, 1, 0, 0, 0);
            DateTime monthEndDate = monthStartDate.AddMonths(1).AddDays(-1).AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);

            MonthStartDate.Set(context, monthStartDate);
            MonthEndDate.Set(context, monthEndDate);
        }
    }
}