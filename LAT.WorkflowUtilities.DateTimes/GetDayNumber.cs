using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class GetDayNumber : WorkFlowActivityBase
    {
        public GetDayNumber() : base(typeof(GetDayNumber)) { }

        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Output("Day Number")]
        public OutArgument<int> DayNumber { get; set; }

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
                GetLocalTime glt = new GetLocalTime();
                int? timeZoneCode = glt.RetrieveTimeZoneCode(localContext.OrganizationService);
                dateToUse = glt.RetrieveLocalTimeFromUtcTime(dateToUse, timeZoneCode, localContext.OrganizationService);
            }

            int dayNumber = dateToUse.Day;

            DayNumber.Set(context, dayNumber);
        }
    }
}