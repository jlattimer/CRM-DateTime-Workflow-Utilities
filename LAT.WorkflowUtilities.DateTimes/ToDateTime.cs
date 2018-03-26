using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class ToDateTime : WorkFlowActivityBase
    {
        public ToDateTime() : base(typeof(ToDateTime)) { }

        [RequiredArgument]
        [Input("Text To Convert")]
        public InArgument<string> TextToConvert { get; set; }

        [Output("Converted Date")]
        public OutArgument<DateTime> ConvertedDate { get; set; }

        [Output("Is Valid Date")]
        public OutArgument<bool> IsValid { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            string textToConvert = TextToConvert.Get(context);

            bool isValid = DateTime.TryParse(textToConvert, out var convertedDate);

            ConvertedDate.Set(context, convertedDate);
            IsValid.Set(context, isValid);
        }
    }
}