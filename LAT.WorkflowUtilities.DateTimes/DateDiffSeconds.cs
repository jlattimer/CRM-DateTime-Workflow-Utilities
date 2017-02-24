using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiffSeconds : CodeActivity
    {
        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [OutputAttribute("Seconds Difference")]
        public OutArgument<int> SecondsDifference { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                DateTime startingDate = StartingDate.Get(executionContext);
                DateTime endingDate = EndingDate.Get(executionContext);

                startingDate = new DateTime(startingDate.Year, startingDate.Month, startingDate.Day, startingDate.Hour,
                    startingDate.Minute, startingDate.Second, startingDate.Kind);

                endingDate = new DateTime(endingDate.Year, endingDate.Month, endingDate.Day, endingDate.Hour,
                    endingDate.Minute, endingDate.Second, endingDate.Kind);

                TimeSpan difference = startingDate - endingDate;

                int secondsDifference = Math.Abs(Convert.ToInt32(difference.TotalSeconds));

                SecondsDifference.Set(executionContext, secondsDifference);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}