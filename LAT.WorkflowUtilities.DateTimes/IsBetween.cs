using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class IsBetween : WorkFlowActivityBase
    {
        public IsBetween() : base(typeof(IsBetween)) { }

        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Date To Validate")]
        public InArgument<DateTime> DateToValidate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [Output("Between")]
        public OutArgument<bool> Between { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime startingDate = StartingDate.Get(context);
            DateTime dateToValidate = DateToValidate.Get(context);
            DateTime endingDate = EndingDate.Get(context);

            var between = dateToValidate > startingDate && dateToValidate < endingDate;

            Between.Set(context, between);
        }
    }
}