using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiffYears : WorkFlowActivityBase
    {
        public DateDiffYears() : base(typeof(DateDiffYears)) { }

        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [Output("Years Difference")]
        public OutArgument<int> YearsDifference { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime startingDate = StartingDate.Get(context);
            DateTime endingDate = EndingDate.Get(context);

            DateTime fromDate;
            DateTime toDate;
            int increment = 0;
            int[] monthDay = { 31, -1, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
            int month;
            int day;

            if (startingDate > endingDate)
            {
                fromDate = endingDate;
                toDate = startingDate;
            }
            else
            {
                fromDate = startingDate;
                toDate = endingDate;
            }

            if (fromDate.Day > toDate.Day)
                increment = monthDay[fromDate.Month - 1];

            if (increment == -1)
                increment = DateTime.IsLeapYear(fromDate.Year) ? 29 : 28;

            if (increment != 0)
            {
                day = toDate.Day + increment - fromDate.Day;
                increment = 1;
            }
            else
                day = toDate.Day - fromDate.Day;

            if (fromDate.Month + increment > toDate.Month)
            {
                month = toDate.Month + 12 - (fromDate.Month + increment);
                increment = 1;
            }
            else
            {
                month = toDate.Month - (fromDate.Month + increment);
                increment = 0;
            }

            int year = toDate.Year - (fromDate.Year + increment);

            int yearsDifference = year;

            YearsDifference.Set(context, yearsDifference);
        }
    }
}