using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Globalization;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class GetMonthName : WorkFlowActivityBase
    {
        public GetMonthName() : base(typeof(GetMonthName)) { }

        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Input("Culture")]
        [Default("en-US")]
        public InArgument<string> Culture { get; set; }

        [Output("Month Name")]
        public OutArgument<string> MonthName { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime dateToUse = DateToUse.Get(context);
            bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(context);
            string culture = Culture.Get(context);

            if (evaluateAsUserLocal)
            {
                int? timeZoneCode = GetLocalTime.RetrieveTimeZoneCode(localContext.OrganizationService);
                dateToUse = GetLocalTime.RetrieveLocalTimeFromUtcTime(dateToUse, timeZoneCode, localContext.OrganizationService);
            }

            string monthName = dateToUse.ToString("MMMM", CultureInfo.CreateSpecificCulture(culture));

            MonthName.Set(context, monthName);
        }
    }
}