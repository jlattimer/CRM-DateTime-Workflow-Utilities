using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Globalization;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class ToUTCString : WorkFlowActivityBase
    {
        public ToUTCString() : base(typeof(ToUTCString)) { }

        [RequiredArgument]
        [Input("Date To Format")]
        public InArgument<DateTime> DateToFormat { get; set; }

        [Input("Culture")]
        [Default("en-US")]
        public InArgument<string> Culture { get; set; }

        [Output("Formatted Date String")]
        public OutArgument<string> FormattedDateString { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime originalDate = DateToFormat.Get(context);
            string cultureIn = Culture.Get(context);

            CultureInfo culture = null;
            if (!string.IsNullOrEmpty(cultureIn))
                culture = new CultureInfo(cultureIn);

            string formattedDateString = originalDate.ToUniversalTime().ToString(culture);

            FormattedDateString.Set(context, formattedDateString);
        }
    }
}