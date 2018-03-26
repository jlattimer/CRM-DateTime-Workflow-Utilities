using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiffSeconds : WorkFlowActivityBase
    {
        public DateDiffSeconds() : base(typeof(DateDiffSeconds)) { }

        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [Output("Seconds Difference")]
        public OutArgument<int> SecondsDifference { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime startingDate = StartingDate.Get(context);
            DateTime endingDate = EndingDate.Get(context);

            startingDate = new DateTime(startingDate.Year, startingDate.Month, startingDate.Day, startingDate.Hour,
                startingDate.Minute, startingDate.Second, startingDate.Kind);

            endingDate = new DateTime(endingDate.Year, endingDate.Month, endingDate.Day, endingDate.Hour,
                endingDate.Minute, endingDate.Second, endingDate.Kind);

            TimeSpan difference = startingDate - endingDate;

            int secondsDifference = Math.Abs(Convert.ToInt32(difference.TotalSeconds));

            SecondsDifference.Set(context, secondsDifference);
        }
    }
}