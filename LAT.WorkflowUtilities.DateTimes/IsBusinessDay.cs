using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class IsBusinessDay : WorkFlowActivityBase
    {
        public IsBusinessDay() : base(typeof(IsBusinessDay)) { }

        [RequiredArgument]
        [Input("Date To Check")]
        public InArgument<DateTime> DateToCheck { get; set; }

        [Input("Holiday/Closure Calendar")]
        [ReferenceTarget("calendar")]
        public InArgument<EntityReference> HolidayClosureCalendar { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Output("Valid Business Day")]
        public OutArgument<bool> ValidBusinessDay { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime dateToCheck = DateToCheck.Get(context);
            bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(context);

            if (evaluateAsUserLocal)
            {
                int? timeZoneCode = GetLocalTime.RetrieveTimeZoneCode(localContext.OrganizationService);
                dateToCheck = GetLocalTime.RetrieveLocalTimeFromUtcTime(dateToCheck, timeZoneCode, localContext.OrganizationService);
            }

            EntityReference holidaySchedule = HolidayClosureCalendar.Get(context);

            bool validBusinessDay = dateToCheck.DayOfWeek != DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday;

            if (!validBusinessDay)
            {
                ValidBusinessDay.Set(context, false);
                return;
            }

            if (holidaySchedule != null)
            {
                Entity calendar = localContext.OrganizationService.Retrieve("calendar", holidaySchedule.Id, new ColumnSet(true));
                if (calendar == null) return;

                EntityCollection calendarRules = calendar.GetAttributeValue<EntityCollection>("calendarrules");
                foreach (Entity calendarRule in calendarRules.Entities)
                {
                    //Date is not stored as UTC
                    DateTime startTime = calendarRule.GetAttributeValue<DateTime>("starttime");

                    //Not same date
                    if (!startTime.Date.Equals(dateToCheck.Date))
                        continue;

                    //Not full day event
                    if (startTime.Subtract(startTime.TimeOfDay) != startTime || calendarRule.GetAttributeValue<int>("duration") != 1440)
                        continue;

                    ValidBusinessDay.Set(context, false);
                    return;
                }
            }

            ValidBusinessDay.Set(context, true);
        }
    }
}