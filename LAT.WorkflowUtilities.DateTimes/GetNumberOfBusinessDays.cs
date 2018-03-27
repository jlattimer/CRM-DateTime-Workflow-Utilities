using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class GetNumberOfBusinessDays : WorkFlowActivityBase
    {
        public GetNumberOfBusinessDays() : base(typeof(GetNumberOfBusinessDays)) { }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

        }
    }
}
