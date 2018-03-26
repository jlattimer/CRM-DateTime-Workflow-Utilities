using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class GetYearStartEnd : WorkFlowActivityBase
    {
        public GetYearStartEnd() : base(typeof(GetYearStartEnd)) { }

        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Output("Year Start Date")]
        public OutArgument<DateTime> YearStartDate { get; set; }

        [Output("Year End Date")]
        public OutArgument<DateTime> YearEndDate { get; set; }

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

            DateTime yearStartDate = new DateTime(dateToUse.Year, 1, 1, 0, 0, 0);
            DateTime yearEndDate = yearStartDate.AddYears(1).AddDays(-1).AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);

            YearStartDate.Set(context, yearStartDate);
            YearEndDate.Set(context, yearEndDate);
        }
    }
}