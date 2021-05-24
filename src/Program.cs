// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using CommandLine.Text;
using CommandLine;
using b2c_ms_graph.Models;
using RateLimiter;
using ComposableAsync;

namespace b2c_ms_graph
{

    public static class Program
    {
        static async Task Main(string[] args)
        {
            string appId = null;
            string appSecret = null;
            string runDecision = null;
            int rateLimit = 0;
            int usersGenerated = 1000;

            Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(o =>
            {
                if (!string.IsNullOrEmpty(o.Application))
                {
                    appId = o.Application;
                }
                if (!string.IsNullOrEmpty(o.Secret))
                {
                    appSecret = o.Secret;
                }
                if (!string.IsNullOrEmpty(o.Decision))
                {
                    runDecision = o.Decision;
                }
                if (o.RateLimit <= 0)
                {
                    rateLimit = 0;
                }
                else
                {
                    rateLimit = o.RateLimit;
                }
                if (o.UserGeneration <= 0)
                {
                    usersGenerated = 0;
                }
                else
                {
                    usersGenerated = o.UserGeneration;
                }
            });

            // Read application settings from appsettings.json (tenant ID, app ID, client secret, etc.)
            AppSettings config = AppSettingsFile.ReadFromJsonFile();
            appId = appId ?? config.AppId;
            appSecret = appSecret ?? config.ClientSecret;

            // Initialize the client credential auth provider
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(appId)
                .WithTenantId(config.TenantId)
                .WithClientSecret(appSecret)
                .Build();
            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            // Set up the Microsoft Graph service client with client credentials
            //GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var clientHandler = new HttpClientHandler
            {
                MaxConnectionsPerServer = 100000
            };

            //var rateHandler = TimeLimiter
            //.GetFromMaxCountByInterval(rateLimit, TimeSpan.FromSeconds(1))
            //.AsDelegatingHandler();

            var handlers = GraphClientFactory.CreateDefaultHandlers(authProvider);
            handlers[0].InnerHandler = clientHandler;
            //handlers.Add(rateHandler);
            var httpClient = GraphClientFactory.Create(handlers);
            GraphServiceClient graphClient = new GraphServiceClient(httpClient);


            PrintCommands();

            try
            {
                while (true)
                {
                    Console.Write("Enter command, then press ENTER: ");
                    string decision = runDecision ?? Console.ReadLine();
                    switch (decision.ToLower())
                    {
                        case "1":
                            await UserService.ListUsers(graphClient);
                            break;
                        case "2":
                            await UserService.GetUserById(graphClient);
                            break;
                        case "3":
                            await UserService.GetUserBySignInName(config, graphClient);
                            break;
                        case "4":
                            await UserService.DeleteUserById(graphClient);
                            break;
                        case "5":
                            await UserService.SetPasswordByUserId(graphClient);
                            break;
                        case "6":
                            await UserService.BulkCreate(config, graphClient);
                            break;
                        case "7":
                            await UserService.BulkCreateTest(config, graphClient, false, usersGenerated, rateLimit);
                            break;
                        case "8":
                            await UserService.BulkCreateTest(config, graphClient, true, usersGenerated, rateLimit);
                            break;
                        case "9":
                            await UserService.CreateUserWithCustomAttribute(graphClient, config.B2cExtensionAppClientId, config.TenantId);
                            break;
                        case "10":
                            await UserService.ListUsersWithCustomAttribute(graphClient, config.B2cExtensionAppClientId);
                            break;
                        //case "11":
                        //    await UserService.UpdateUserTest(config, graphClient);
                        //    break;
                        case "help":
                            Program.PrintCommands();
                            break;
                        case "exit":
                            return;
                        default:
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Invalid command. Enter 'help' to show a list of commands.");
                            Console.ResetColor();
                            break;
                    }

                    runDecision = null;
                    Console.ResetColor();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex}");
                Console.ResetColor();
            }
            Console.ReadLine();
        }

        private static void PrintCommands()
        {
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine("Command  Description");
            Console.WriteLine("====================");
            Console.WriteLine("[1]      Get all users (one page)");
            Console.WriteLine("[2]      Get user by object ID");
            Console.WriteLine("[3]      Get user by sign-in name");
            Console.WriteLine("[4]      Delete user by object ID");
            Console.WriteLine("[5]      Update user password");
            Console.WriteLine("[6]      Create users (bulk import csv)");
            Console.WriteLine("[7]      Create Random users (bulk import test)");
            Console.WriteLine("[8]      Create Random users Batch (bulk import batch test)");
            Console.WriteLine("[9]      Create user with custom attributes and show result");
            Console.WriteLine("[10]     Get all users (one page) with custom attributes");
            //Console.WriteLine("[11]     Update user test (updates top 1000 users with new random properties)");
            Console.WriteLine("[help]   Show available commands");
            Console.WriteLine("[exit]   Exit the program");
            Console.WriteLine("-------------------------");
        }
    }
}
