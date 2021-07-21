using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace MicrosoftGraphSample.ConsoleApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var clientId = "xxx";
            var secret = "xxx";
            var tenantId = "xxx";

            // GraphServiceClient を生成します。
            var confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(secret)
                .WithTenantId(tenantId)
                .Build();

            // 以下のパターンを利用する場合は、Microsoft.Graph.Auth が必要です。
            // ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);
            // GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async req =>
            {
                var scopes = new[] { @"https://graph.microsoft.com//.default" };
                var authenticationResult = await confidentialClientApplication.AcquireTokenForClient(scopes).ExecuteAsync();
                req.Headers.Authorization = new AuthenticationHeaderValue("bearer", authenticationResult.AccessToken);
                req.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");
            }));


            // ユーザ情報を取得するサンプルです。
            var users = await graphClient.Users
                .Request()
                .Select("id,displayName,mail,jobTitle,department,officeLocation")
                .GetAsync();

            foreach (var user in users)
            {
                Console.WriteLine($"{user.Id} - DisplayName: {user.DisplayName} (Email: {user.Mail}, Title: {user.JobTitle}, Department: {user.Department}, Location: {user.OfficeLocation})");

                var queryOptions = new List<QueryOption>
                {
                    new QueryOption("startDateTime", "2021-07-01T00:00:00-00:00"),
                    new QueryOption("endDateTime", "2021-08-01T00:00:00.00-00:00")
                };

                // カレンダー情報を取得するサンプルです。
                var calendarView = await graphClient
                    .Users[user.Id]
                    .CalendarView
                    .Request(queryOptions)
                    // GlobalObjectId (PidLidGlobalObjectId) を取得する場合、拡張プロパティを設定します。
                    // https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/dff9f123-fb9b-4581-a853-e70dd0cda6a7
                    .Expand("singleValueExtendedProperties($filter=id eq 'Binary {6ED8DA90-450B-101B-98DA-00AA003F1305} Id 0x0003')")
                    .GetAsync();

                foreach (var c in calendarView)
                {
                    Console.WriteLine($"{c.Id} {c.Subject} {c.Start.DateTime} {c.ResponseStatus.Response}");

                    // GlobalObjectId
                    Console.WriteLine(c.SingleValueExtendedProperties.First().Value);

                    // ICalUId から GlobalObjectId を取得する場合の例です。
                    Console.WriteLine(GetObjectIdStringFromUid(c.ICalUId));
                }
            }

            // 条件を指定してユーザ情報を取得するサンプルです。
            users = await graphClient.Users
                .Request()
                .Select("id,displayName,mail,jobTitle,department,officeLocation")
                .Filter("department eq 'テスト部署' and jobTitle eq 'てすと役職'")
                .GetAsync();

            foreach (var user in users)
            {
                Console.WriteLine(
                    $"{user.Id} - DisplayName: {user.DisplayName} (Email: {user.Mail}, Title: {user.JobTitle}, Department: {user.Department}, Location: {user.OfficeLocation})");
            }

            // イベントを登録します。
            var @event = new Event
            {
                Subject = "Let's go for lunch",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "Does mid month work for you?"
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = "2021-07-20T12:00:00",
                    TimeZone = "Tokyo Standard Time"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = "2021-07-20T14:00:00",
                    TimeZone = "Tokyo Standard Time"
                },
                Location = new Location
                {
                    DisplayName = "Harry's Bar"
                },
                Attendees = new List<Attendee>
                {
                    new Attendee
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "yoshiyuki.tsunomoto@wasiwasimaru.onmicrosoft.com"
                        },
                        Type = AttendeeType.Required
                    }
                },
                TransactionId = "7E163156-7762-4BEB-A1C6-729EA81755A7"
            };

            var id = "AdeleV@wasiwasimaru.onmicrosoft.com";
            await graphClient.Users[id]
                .Calendar
                .Events
                .Request()
                .AddAsync(@event);

            Console.ReadKey();
        }

        // 
        /// <summary>
        /// ICalUId から GlobalObjectId を取得します。
        /// </summary>
        /// <remarks>
        /// https://social.msdn.microsoft.com/Forums/sqlserver/en-US/e1714a15-1ef7-4868-a701-f53c5ceae1ad/ews-api-differences-in-icaluid-returned-when-appointments-are-created-by-office-365-account-vs?forum=exchangesvrdevelopment
        /// </remarks>
        /// <param name="id"></param>
        /// <returns></returns>
        private static string GetObjectIdStringFromUid(string id)
        {
            var buffer = new byte[id.Length / 2];
            for (var i = 0; i < id.Length / 2; i++)
            {
                var hexValue = byte.Parse(id.Substring(i * 2, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
                buffer[i] = hexValue;
            }
            return Convert.ToBase64String(buffer);
        }
    }
}
