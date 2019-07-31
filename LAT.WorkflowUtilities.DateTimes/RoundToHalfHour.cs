using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class RoundToHalfHour : WorkFlowActivityBase
    {
        public RoundToHalfHour() : base(typeof(DateDiffMinutes)) { }

        [RequiredArgument]
        [Input("Date To Round")]
        public InArgument<DateTime> DateToRound { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Round Down (Otherwise Round Up)?")]
        public InArgument<bool> RoundDown { get; set; }

        [Output("Rounded Date")]
        public OutArgument<DateTime> RoundedDate { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            var dateToRound = DateToRound.Get(context);
            var roundDown = RoundDown.Get(context);

            dateToRound = dateToRound.AddSeconds(-dateToRound.Second);
            var roundedDate = dateToRound;
            var minute = dateToRound.Minute;
            if (minute > 0 && minute < 30)
            {
                roundedDate = roundDown
                    ? dateToRound.AddMinutes(-minute) // 0
                    : dateToRound.AddMinutes(30 - minute); // 30
            }
            else if (minute > 30 && minute < 60)
            {
                roundedDate = roundDown
                    ? dateToRound.AddMinutes(-(minute - 30)) // 60
                    : dateToRound.AddMinutes(60 - minute); // Next hour
            }

            RoundedDate.Set(context, roundedDate);
        }
    }
}