using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiffHours : WorkFlowActivityBase
    {
        public DateDiffHours() : base(typeof(DateDiffHours)) { }

        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [Output("Hours Difference")]
        public OutArgument<int> HoursDifference { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime startingDate = StartingDate.Get(context);
            DateTime endingDate = EndingDate.Get(context);

            TimeSpan difference = startingDate - endingDate;

            int hoursDifference = Math.Abs(Convert.ToInt32(difference.TotalHours));

            HoursDifference.Set(context, hoursDifference);
        }
    }
}