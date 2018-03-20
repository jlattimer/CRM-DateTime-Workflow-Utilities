using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class IsSameDay : WorkFlowActivityBase
    {
        public IsSameDay() : base(typeof(IsSameDay)) { }

        [RequiredArgument]
        [Input("First Date")]
        public InArgument<DateTime> FirstDate { get; set; }

        [RequiredArgument]
        [Input("Second Date")]
        public InArgument<DateTime> SecondDate { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Output("Same Day")]
        public OutArgument<bool> SameDay { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime firstDate = FirstDate.Get(context);
            DateTime secondDate = SecondDate.Get(context);
            bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(context);

            if (evaluateAsUserLocal)
            {
                GetLocalTime glt = new GetLocalTime();
                int? timeZoneCode = glt.RetrieveTimeZoneCode(localContext.OrganizationService);
                firstDate = glt.RetrieveLocalTimeFromUtcTime(firstDate, timeZoneCode, localContext.OrganizationService);
                secondDate = glt.RetrieveLocalTimeFromUtcTime(secondDate, timeZoneCode, localContext.OrganizationService);
            }

            bool sameDay = firstDate.Date == secondDate.Date;

            SameDay.Set(context, sameDay);
        }
    }
}