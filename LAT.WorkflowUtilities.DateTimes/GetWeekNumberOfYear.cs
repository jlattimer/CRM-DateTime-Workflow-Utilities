using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Globalization;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class GetWeekNumberOfYear : WorkFlowActivityBase
    {
        public GetWeekNumberOfYear() : base(typeof(GetWeekNumberOfYear)) { }

        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Output("Week Number Of Year")]
        public OutArgument<int> WeekNumberOfYear { get; set; }

        //http://blogs.msdn.com/b/shawnste/archive/2006/01/24/iso-8601-week-of-year-format-in-microsoft-net.aspx

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime dateToUse = DateToUse.Get(context);
            bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(context);

            if (evaluateAsUserLocal)
            {
                GetLocalTime glt = new GetLocalTime();
                int? timeZoneCode = glt.RetrieveTimeZoneCode(localContext.OrganizationService);
                dateToUse = glt.RetrieveLocalTimeFromUtcTime(dateToUse, timeZoneCode, localContext.OrganizationService);
            }

            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;

            int weekNumberOfYear = cal.GetWeekOfYear(dateToUse, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);

            WeekNumberOfYear.Set(context, weekNumberOfYear);
        }
    }
}