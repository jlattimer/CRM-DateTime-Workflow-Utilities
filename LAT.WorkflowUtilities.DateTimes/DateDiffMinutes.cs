using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiffMinutes : WorkFlowActivityBase
    {
        public DateDiffMinutes() : base(typeof(DateDiffMinutes)) { }

        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [Output("Minutes Difference")]
        public OutArgument<int> MinutesDifference { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime startingDate = StartingDate.Get(context);
            DateTime endingDate = EndingDate.Get(context);

            startingDate = new DateTime(startingDate.Year, startingDate.Month, startingDate.Day, startingDate.Hour,
                startingDate.Minute, 0, startingDate.Kind);

            endingDate = new DateTime(endingDate.Year, endingDate.Month, endingDate.Day, endingDate.Hour,
                endingDate.Minute, 0, endingDate.Kind);

            TimeSpan difference = startingDate - endingDate;

            int minutesDifference = Math.Abs(Convert.ToInt32(difference.TotalMinutes));

            MinutesDifference.Set(context, minutesDifference);
        }
    }
}