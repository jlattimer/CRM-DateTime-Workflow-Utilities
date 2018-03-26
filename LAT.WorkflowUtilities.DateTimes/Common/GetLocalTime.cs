using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace LAT.WorkflowUtilities.DateTimes.Common
{
    public class GetLocalTime
    {
        public static int? RetrieveTimeZoneCode(IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions = {
                            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)
                        }
                    }
                });

            var e = currentUserSettings.Entities[0].ToEntity<Entity>();

            return (int?)e.Attributes["timezonecode"];
        }

        public static DateTime RetrieveLocalTimeFromUtcTime(DateTime utcTime, int? timeZoneCode, IOrganizationService service)
        {
            if (!timeZoneCode.HasValue)
                return DateTime.Now;

            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };

            OrganizationResponse response = service.Execute(request);
            return (DateTime)response.Results["LocalTime"];
        }
    }
}