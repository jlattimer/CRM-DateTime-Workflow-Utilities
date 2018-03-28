using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace LAT.WorkflowUtilities.DateTimes.Common
{
    /// <summary>
    /// Class containing common functionality relating to business days
    /// </summary>
    public static class BusinessDayLogic
    {
        public static readonly int MinutesInDay = 1440;

        /// <summary>
        /// Method to check if a day is a business day (i.e. Monday - Friday and not a holiday)
        /// </summary>
        /// <param name="dateToCheck">Date to evaluate</param>
        /// <param name="service">Organisation Service to query for calendar info</param>
        /// <param name="holidaySchedule">The calendar to use for closures - if passed the service must be passed too</param>
        /// <returns>True if the day is a business day</returns>
        public static bool IsBusinessDay(this DateTime dateToCheck, IOrganizationService service = null, EntityReference holidaySchedule = null)
        {
            var validBusinessDay = dateToCheck.DayOfWeek != DayOfWeek.Saturday && dateToCheck.DayOfWeek != DayOfWeek.Sunday;

            if (!validBusinessDay)
            {
                return false;
            }

            if (holidaySchedule == null)
            {
                return true;
            }

            if (service == null)
            {
                throw new ArgumentException($"{holidaySchedule} is provided but the Organisation Service is null.", nameof(service));
            }

            var calendar = service.Retrieve("calendar", holidaySchedule.Id, new ColumnSet(true));
            if (calendar == null)
            {
                return true;
            }

            var calendarRules = calendar.GetAttributeValue<EntityCollection>("calendarrules");

            foreach (var calendarRule in calendarRules.Entities)
            {
                // Date is not stored as UTC
                var startTime = calendarRule.GetAttributeValue<DateTime>("starttime");
                var duration = calendarRule.GetAttributeValue<int>("duration");
                var endTime = startTime.AddMinutes(duration);

                if (dateToCheck.IsBetween(startTime, endTime))
                {
                    // Date is covered by the holiday - check to see if the holiday is a full day
                    if (duration < MinutesInDay)
                    {
                        // Closure is not a full day, ignore
                        continue;
                    }

                    if (dateToCheck.Date == endTime.Date)
                    {
                        // The date to check matches the end bound which means either the calendar rules doesn't actually cover it (midnight overlap) or this ends part way through the day, so not a full closure
                        continue;
                    }

                    if (dateToCheck.Date == startTime.Date && !(startTime.Hour == 0 && startTime.Minute == 0 && startTime.Second == 0))
                    {
                        // The date to check is the start day which doesnt start at 00:00:00, therefore this is only a partial day
                        continue;
                    }

                    // Looks like a full closure at this point, this is not a business day
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Extension method to check if a point in time falls between 2 others (inclusive)
        /// </summary>
        /// <param name="input">Time to check</param>
        /// <param name="date1">Start of the time window</param>
        /// <param name="date2">End of the time window</param>
        /// <returns>True if the input is in the range or on the bounds</returns>
        public static bool IsBetween(this DateTime input, DateTime date1, DateTime date2)
        {
            return input >= date1 && input <= date2;
        }
    }
}