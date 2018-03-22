using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class SetTime : WorkFlowActivityBase
    {
        public SetTime() : base(typeof(SetTime)) { }

        [RequiredArgument]
        [Input("Date To Update")]
        public InArgument<DateTime> DateToUpdate { get; set; }

        [RequiredArgument]
        [Output("Hour (24-hour)")]
        [Default("12")]
        public InArgument<int> Hour { get; set; }

        [RequiredArgument]
        [Output("Minute")]
        [Default("0")]
        public InArgument<int> Minute { get; set; }

        [RequiredArgument]
        [Output("Second")]
        [Default("0")]
        public InArgument<int> Second { get; set; }

        [Output("Updated Date")]
        public OutArgument<DateTime> UpdatedDate { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime dateToUpdate = DateToUpdate.Get(context);
            int hour = Hour.Get(context);
            if (hour > 23 || hour < 0)
                throw new InvalidPluginExecutionException("Invalid hour");

            int minute = Minute.Get(context);
            if (minute > 59 || minute < 0)
                throw new InvalidPluginExecutionException("Invalid minute");

            int second = Second.Get(context);
            if (second > 59 || second < 0)
                throw new InvalidPluginExecutionException("Invalid second");

            DateTime updatedDate = new DateTime(dateToUpdate.Year, dateToUpdate.Month, dateToUpdate.Day, hour, minute, second);

            UpdatedDate.Set(context, updatedDate);
        }
    }
}