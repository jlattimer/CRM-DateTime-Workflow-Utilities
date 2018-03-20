using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class AddYears : WorkFlowActivityBase
    {
        public AddYears() : base(typeof(AddYears)) { }

        [RequiredArgument]
        [Input("Original Date")]
        public InArgument<DateTime> OriginalDate { get; set; }

        [RequiredArgument]
        [Input("Years To Add")]
        public InArgument<int> YearsToAdd { get; set; }

        [Output("Updated Date")]
        public OutArgument<DateTime> UpdatedDate { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime originalDate = OriginalDate.Get(context);
            int yearsToAdd = YearsToAdd.Get(context);

            DateTime updatedDate = originalDate.AddYears(yearsToAdd);

            UpdatedDate.Set(context, updatedDate);
        }
    }
}