using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiffDays : WorkFlowActivityBase
    {
        public DateDiffDays() : base(typeof(DateDiffDays)) { }

        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [Output("Days Difference")]
        public OutArgument<int> DaysDifference { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime startingDate = StartingDate.Get(context);
            DateTime endingDate = EndingDate.Get(context);

            TimeSpan difference = startingDate - endingDate;

            int daysDifference = Math.Abs(Convert.ToInt32(difference.TotalDays));

            DaysDifference.Set(context, daysDifference);
        }
    }
}