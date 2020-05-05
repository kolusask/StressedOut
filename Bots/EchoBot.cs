// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Net.Http;
using System.IO;
using System.Text;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class User
    {
        public string Name{get;set;}
        public bool IsAdmin {get;set;}
        public bool IsManager {get;set;}
        public List<User> Employees {get;set;} = new List<User>();
    }
    public class Issue
    {
        public string Name{get;set;}
        public int Priority{get;set;}
        public DateTime StartTime {get;set;}
        public DateTime EndTime {get;set;}
    }
    public class Meetings
    {
        public string Topic {get;set;}
        public string Description {get;set;}
        public List<User> Attendees{get;set;} = new List<User>();
        public DateTime StartTime{get;set;}
        public DateTime EndTime{get;set;}
    }

    public class EchoBot : ActivityHandler
    {
        List<User> AllUsers{get;set;} = new List<User>();
        private void AddUser(User user)
        {
            AllUsers.Add(user);
            //serialize JSON
            var tmp = JsonConvert.SerializeObject(AllUsers);
            using (System.IO.FileStream fs = System.IO.File.Open("users.json", FileMode.OpenOrCreate))
            {
                var byteArray = new UTF8Encoding(true).GetBytes(tmp);
                fs.Write(byteArray, 0, byteArray.Length);
            }
        }

        //private User GetUser()
        //{
            
        //}

        public bool IsAvailable(){
            return (Math.Pow(DateTime.Now.Day, DateTime.Now.Hour) * DateTime.Now.Minute + DateTime.Now.Second) %2 == 0;
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var query = turnContext.Activity.Text;
            var tokens = query.Split('|');
            var func = tokens[0];
            var body = query.Substring(func.Length);
            
            if(func == "AddUser"){
                var name = tokens[1];
                var role = tokens[2];
                var isAdmin = role == "Admin";
                var isManager = role == "Manager";

                var user = new User
                {
                    Name = name,
                    IsAdmin = isAdmin,
                    IsManager = isManager
                };

                AddUser(user);
            }
            
            if (func == "AddEvent")
            {
                string summary = tokens[1];
                string description = tokens[2];
                string location = tokens[3];
                string startTime = tokens[4];
                string startTZone = tokens[5];
                string endTime = tokens[6];
                string endTZone = tokens[7];
                List<string> emails = new List<string>();
                for (int i = 8; i < tokens.Length; i++)
                    emails.Add(tokens[i]);

                IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
                    .Create("d81502a3-9729-448d-938b-e8b2dcccd437")
                    .Build();

                string[] scopes = { "https://graph.microsoft.com/Calendars.ReadWrite" };
                // Create an authentication provider by passing in a client application and graph scopes.
                DeviceCodeProvider authProvider = new DeviceCodeProvider(publicClientApplication, scopes);
                GraphServiceClient graphClient = new GraphServiceClient( authProvider );

                List<Attendee> attendees = new List<Attendee>{};
                foreach(string s in emails)
                {
                    if(!IsAvailable())
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Text($"{s} is busy", $"{s} is busy"), cancellationToken);
                        return;
                    }
                    attendees.Add(
                        new Attendee
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = s,
                                Name = ""
                            },
                            Type = AttendeeType.Required
                        }
                    );
                }
                

                var @event = new Event
                {
                    Subject = summary,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = description
                    },
                    Start = new DateTimeTimeZone
                    {
                        DateTime = startTime,
                        TimeZone = startTZone
                    },
                    End = new DateTimeTimeZone
                    {
                        DateTime = endTime,
                        TimeZone = endTZone
                    },
                    Location = new Location
                    {
                        DisplayName = location
                    },
                    Attendees = attendees
                };

                var res = await graphClient.Me.Calendar.Events
                    .Request()
                    .AddAsync(@event);
            }
            /*
            else if(func == "IsAvailable")
            {
                
                using(var client = new HttpClient())
                {
                    var url = "API_URL" + "?name=" + body;
                    //get about extraversion and hours
                    var response = await client.GetAsync(url);
                    if(response.IsSuccessStatusCode){
                        await turnContext.SendActivityAsync(MessageFactory.Text("Available", "Available"), cancellationToken);
                    }
                }
            }
            */
            
            await turnContext.SendActivityAsync(MessageFactory.Text(query, query), cancellationToken);
            //return new Task();
        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var welcomeText = "Hello and welcome!";
            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText, welcomeText), cancellationToken);
                }
            }
        }
    }
}
