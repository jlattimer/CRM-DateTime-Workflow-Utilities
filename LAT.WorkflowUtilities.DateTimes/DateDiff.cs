using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Text;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiff : WorkFlowActivityBase
    {
        public DateDiff() : base(typeof(DateDiff)) { }

        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Show Seconds")]
        public InArgument<bool> ShowSeconds { get; set; }

        [Output("Difference")]
        public OutArgument<string> Difference { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime startingDate = StartingDate.Get(context);
            DateTime endingDate = EndingDate.Get(context);
            bool showSeconds = ShowSeconds.Get(context);

            TimeSpan diff = startingDate - endingDate;

            StringBuilder sb = new StringBuilder();

            //if (diff.Days > 0)
                sb.Append($"{diff.Days}d.");

            //if (diff.Hours > 0)
                sb.Append($"{diff.Hours:00}h:");

            //if (diff.Minutes > 0)
                sb.Append($"{diff.Minutes:00}m");

            //if (showSeconds && diff.Seconds > 0)
            if (showSeconds)
                sb.Append($":{diff.Seconds:00}s");

            Difference.Set(context, sb.ToString());
        }
    }
}