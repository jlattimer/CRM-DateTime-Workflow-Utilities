using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class GetDayNumberOfWeek : WorkFlowActivityBase
    {
        public GetDayNumberOfWeek() : base(typeof(GetDayNumberOfWeek)) { }

        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Output("Day Number Of Week")]
        public OutArgument<int> DayNumberOfWeek { get; set; }

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

            int dayNumberOfWeek = (int)dateToUse.DayOfWeek + 1;

            DayNumberOfWeek.Set(context, dayNumberOfWeek);
        }
    }
}