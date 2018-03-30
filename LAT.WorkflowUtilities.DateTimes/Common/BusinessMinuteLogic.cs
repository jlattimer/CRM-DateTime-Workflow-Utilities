using Microsoft.Xrm.Sdk;
using System;

namespace LAT.WorkflowUtilities.DateTimes.Common
{
    /// <summary>
    /// Class containing common functionality relating to business minutes
    /// </summary>
    public static class BusinessMinuteLogic
    {
        /// <summary>
        /// Method to check if a minute is part of a business day (i.e. Monday - Friday and not a holiday)
        /// </summary>
        /// <param name="dateToCheck">Date to evaluate</param>
        /// <param name="calendar">The holiday/closure calendar to use</param>
        /// <returns>True if the day is a business day</returns>
        public static bool IsBusinessMinute(this DateTime dateToCheck, Entity calendar)
        {
            var validBusinessDay = dateToCheck.DayOfWeek != DayOfWeek.Saturday && dateToCheck.DayOfWeek != DayOfWeek.Sunday;

            if (!validBusinessDay)
                return false;

            if (calendar == null)
                return true;

            var calendarRules = calendar.GetAttributeValue<EntityCollection>("calendarrules");

            foreach (var calendarRule in calendarRules.Entities)
            {
                // Date is not stored as UTC
                var startTime = calendarRule.GetAttributeValue<DateTime>("starttime");
                // Subtract 1 so the last minute is not double counted 
                // 4/2/2018 12:00 AM + 1440 minutes = 4/3/2018 12:00 AM - so the last minute is handled twice
                var duration = calendarRule.GetAttributeValue<int>("duration") - 1;
                var endTime = startTime.AddMinutes(duration);

                // Date is during a holiday
                if (dateToCheck.IsBetween(startTime, endTime))
                    return false;
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
        private static bool IsBetween(this DateTime input, DateTime date1, DateTime date2)
        {
            return input >= date1 && input <= date2;
        }
    }
}