using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class AddBusinessDays : WorkFlowActivityBase
    {
        public AddBusinessDays() : base(typeof(AddBusinessDays)) { }

        [RequiredArgument]
        [Input("Original Date")]
        public InArgument<DateTime> OriginalDate { get; set; }

        [RequiredArgument]
        [Input("Business Days To Add")]
        public InArgument<int> BusinessDaysToAdd { get; set; }

        [Input("Holiday/Closure Calendar")]
        [ReferenceTarget("calendar")]
        public InArgument<EntityReference> HolidayClosureCalendar { get; set; }

        [Output("Updated Date")]
        public OutArgument<DateTime> UpdatedDate { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            DateTime originalDate = OriginalDate.Get(context);
            int businessDaysToAdd = BusinessDaysToAdd.Get(context);
            EntityReference holidaySchedule = HolidayClosureCalendar.Get(context);

            Entity calendar = null;
            EntityCollection calendarRules = null;
            if (holidaySchedule != null)
            {
                calendar = localContext.OrganizationService.Retrieve("calendar", holidaySchedule.Id, new ColumnSet(true));
                if (calendar != null)
                    calendarRules = calendar.GetAttributeValue<EntityCollection>("calendarrules");
            }

            DateTime tempDate = originalDate;

            if (businessDaysToAdd > 0)
            {
                while (businessDaysToAdd > 0)
                {
                    tempDate = tempDate.AddDays(1);
                    if (tempDate.DayOfWeek == DayOfWeek.Sunday || tempDate.DayOfWeek == DayOfWeek.Saturday)
                        continue;

                    if (calendar == null)
                    {
                        businessDaysToAdd--;
                        continue;
                    }

                    bool isHoliday = false;
                    foreach (Entity calendarRule in calendarRules.Entities)
                    {
                        DateTime startTime = calendarRule.GetAttributeValue<DateTime>("starttime");

                        //Not same date
                        if (!startTime.Date.Equals(tempDate.Date))
                            continue;

                        //Not full day event
                        if (startTime.Subtract(startTime.TimeOfDay) != startTime || calendarRule.GetAttributeValue<int>("duration") != 1440)
                            continue;

                        isHoliday = true;
                        break;
                    }
                    if (!isHoliday)
                        businessDaysToAdd--;
                }
            }
            else if (businessDaysToAdd < 0)
            {
                while (businessDaysToAdd < 0)
                {
                    tempDate = tempDate.AddDays(-1);
                    if (tempDate.DayOfWeek == DayOfWeek.Sunday || tempDate.DayOfWeek == DayOfWeek.Saturday)
                        continue;

                    if (calendar == null)
                    {
                        businessDaysToAdd++;
                        continue;
                    }

                    bool isHoliday = false;
                    foreach (Entity calendarRule in calendarRules.Entities)
                    {
                        DateTime startTime = calendarRule.GetAttributeValue<DateTime>("starttime");

                        //Not same date
                        if (!startTime.Date.Equals(tempDate.Date))
                            continue;

                        //Not full day event
                        if (startTime.Subtract(startTime.TimeOfDay) != startTime || calendarRule.GetAttributeValue<int>("duration") != 1440)
                            continue;

                        isHoliday = true;
                        break;
                    }
                    if (!isHoliday)
                        businessDaysToAdd++;
                }
            }

            DateTime updatedDate = tempDate;

            UpdatedDate.Set(context, updatedDate);
        }
    }
}