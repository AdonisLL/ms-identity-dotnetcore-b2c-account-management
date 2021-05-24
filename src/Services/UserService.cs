// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Newtonsoft.Json;
using RandomNameGen;
using b2c_ms_graph.Helpers;
using System.Net.Http;
using System.Text;
using RateLimiter;
using ComposableAsync;
using Newtonsoft.Json.Linq;

namespace b2c_ms_graph
{
    public class UserService
    {
        public static List<User> failedUserList = new List<User>();


        private long GetElapsedTimeInTicks(int elapsedTimeMs)
        {
            return elapsedTimeMs * TimeSpan.TicksPerMillisecond;
        }

        /*
         * Calculates the minimum elapsed time required for a request to conform, assuming the previous window met the request limit
         * This equation is based on the one used by SlidingWindow
         */
        private long GetMinimumElapsedTimeInTicks(int requestLimit, int requestIntervalMs, int currentRequestCount)
        {
            return GetElapsedTimeInTicks(-1 * (requestIntervalMs * (requestLimit - currentRequestCount - 1) / requestLimit - requestIntervalMs));
        }


        public static async Task ListUsers(GraphServiceClient graphClient)
        {
            Console.WriteLine("Getting list of users...");

            // Get all users (one page)
            var result = await graphClient.Users
                .Request()
                .Select(e => new
                {
                    e.DisplayName,
                    e.Id,
                    e.Identities
                })
                .GetAsync();

            foreach (var user in result.CurrentPage)
            {
                Console.WriteLine(JsonConvert.SerializeObject(user));
            }
        }

        public static async Task ListUsersWithCustomAttribute(GraphServiceClient graphClient, string b2cExtensionAppClientId)
        {
            if (string.IsNullOrWhiteSpace(b2cExtensionAppClientId))
            {
                throw new ArgumentException("B2cExtensionAppClientId (its Application ID) is missing from appsettings.json. Find it in the App registrations pane in the Azure portal. The app registration has the name 'b2c-extensions-app. Do not modify. Used by AADB2C for storing user data.'.", nameof(b2cExtensionAppClientId));
            }

            // Declare the names of the custom attributes
            const string customAttributeName1 = "FavouriteSeason";
            const string customAttributeName2 = "LovesPets";

            // Get the complete name of the custom attribute (Azure AD extension)
            Helpers.B2cCustomAttributeHelper helper = new Helpers.B2cCustomAttributeHelper(b2cExtensionAppClientId);
            string favouriteSeasonAttributeName = helper.GetCompleteAttributeName(customAttributeName1);
            string lovesPetsAttributeName = helper.GetCompleteAttributeName(customAttributeName2);

            Console.WriteLine($"Getting list of users with the custom attributes '{customAttributeName1}' (string) and '{customAttributeName2}' (boolean)");
            Console.WriteLine();

            // Get all users (one page)
            var result = await graphClient.Users
                .Request()
                .Select($"id,displayName,identities,{favouriteSeasonAttributeName},{lovesPetsAttributeName}")
                .GetAsync();

            foreach (var user in result.CurrentPage)
            {
                Console.WriteLine(JsonConvert.SerializeObject(user));

                // Only output the custom attributes...
                //Console.WriteLine(JsonConvert.SerializeObject(user.AdditionalData));
            }
        }

        public static async Task GetUserById(GraphServiceClient graphClient)
        {
            Console.Write("Enter user object ID: ");
            string userId = Console.ReadLine();

            Console.WriteLine($"Looking for user with object ID '{userId}'...");

            try
            {
                // Get user by object ID
                var result = await graphClient.Users[userId]
                    .Request()
                    .Select(e => new
                    {
                        e.DisplayName,
                        e.Id,
                        e.Identities
                    })
                    .GetAsync();

                if (result != null)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(result));
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }

        public static async Task GetUserBySignInName(AppSettings config, GraphServiceClient graphClient)
        {
            Console.Write("Enter user sign-in name (username or email address): ");
            string userId = Console.ReadLine();

            Console.WriteLine($"Looking for user with sign-in name '{userId}'...");

            try
            {
                // Get user by sign-in name
                var result = await graphClient.Users
                    .Request()
                    .Filter($"identities/any(c:c/issuerAssignedId eq '{userId}' and c/issuer eq '{config.TenantId}')")
                    .Select(e => new
                    {
                        e.DisplayName,
                        e.Id,
                        e.Identities
                    })
                    .GetAsync();

                if (result != null)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(result));
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }

        public static async Task UpdateUserTest(AppSettings config, GraphServiceClient graphClient)
        {

            try
            {
                List<User> updatedUsers = new List<User>();

                var tasks = new List<Task>();
                var watch = System.Diagnostics.Stopwatch.StartNew();

                var result = await graphClient.Users
                    .Request()
                    .Top(999)
                    .Select(e => new
                    {
                        e.DisplayName,
                        e.Id,
                        e.Identities
                    })
                    .GetAsync();


                foreach (var u in result)
                {
                    UpdateTestUserProperties(u);
                    updatedUsers.Add(u);
                }

                var timeConstraint = TimeLimiter.GetFromMaxCountByInterval(1000, TimeSpan.FromSeconds(1));
                foreach (var u in updatedUsers)
                {
                    await timeConstraint;
                    tasks.Add(UpdateGraphUser(graphClient, u));
                }


                Task allTasks = Task.WhenAll(tasks);
                try
                {
                    await allTasks;
                }
                catch
                {
                    AggregateException allExceptions = allTasks.Exception;
                }

                watch.Stop();
                var elapsedMs = watch.ElapsedMilliseconds;
                Console.WriteLine($"Update Completed in {TimeSpan.FromMilliseconds(elapsedMs).TotalSeconds} Seconds");

            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }

        public static async Task DeleteUserById(GraphServiceClient graphClient)
        {
            Console.Write("Enter user object ID: ");
            string userId = Console.ReadLine();

            Console.WriteLine($"Looking for user with object ID '{userId}'...");

            try
            {
                // Delete user by object ID
                await graphClient.Users[userId]
                   .Request()
                   .DeleteAsync();

                Console.WriteLine($"User with object ID '{userId}' successfully deleted.");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }

        public static async Task SetPasswordByUserId(GraphServiceClient graphClient)
        {
            Console.Write("Enter user object ID: ");
            string userId = Console.ReadLine();

            Console.Write("Enter new password: ");
            string password = Console.ReadLine();

            Console.WriteLine($"Looking for user with object ID '{userId}'...");

            var user = new User
            {
                PasswordPolicies = "DisablePasswordExpiration,DisableStrongPassword",
                PasswordProfile = new PasswordProfile
                {
                    ForceChangePasswordNextSignIn = false,
                    Password = password,
                }
            };

            try
            {
                // Update user by object ID
                await graphClient.Users[userId]
                   .Request()
                   .UpdateAsync(user);

                Console.WriteLine($"User with object ID '{userId}' successfully updated.");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }


        public static async Task BulkCreateTest(AppSettings config, GraphServiceClient graphClient,
            bool isBatch = false, int GeneratedUserCount = 0, int RequestPerMinute = 0)
        {
            try
            {
                if (GeneratedUserCount == 0)
                {
                    Console.WriteLine("How many test users should be generated?");
                    GeneratedUserCount = int.Parse(Console.ReadLine());
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Should be an integer, Defaulting to 1000 users");
                GeneratedUserCount = 1000;
            }

            try
            {
                if (RequestPerMinute == 0)
                {
                    Console.WriteLine("Maximum request per minute??");
                    RequestPerMinute = int.Parse(Console.ReadLine());
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Should be an integer, Defaulting to 300 request per minute");
                RequestPerMinute = 300;
            }

            Console.WriteLine("Starting bulk create operation...");
            var userList = GenerateRandomGraphUsers(GeneratedUserCount);
            userList.ToList().ForEach(x => x.SetB2CProfile(config.TenantId));

            var batchSize = 20;
            int numberOfBatches = (int)Math.Ceiling((double)userList.Count / batchSize);

            var tasks = new List<Task>();
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var timeConstraint = TimeLimiter.GetFromMaxCountByInterval(RequestPerMinute, TimeSpan.FromSeconds(1));

            if (isBatch)
            {
                for (int i = 0; i < numberOfBatches; i++)
                {

                    var items = userList.Skip(i * batchSize).Take(batchSize).ToList();
                    await timeConstraint;
                    tasks.Add(CreateRandomGraphUserBatch(items, graphClient));

                }
            }
            else
            {

                foreach (var u in userList)
                {
                    await timeConstraint;
                    tasks.Add(AddGraphUser(graphClient, u));
                }

            }

            Task allTasks = Task.WhenAll(tasks);
            try
            {
                await allTasks;
            }
            catch
            {
                AggregateException allExceptions = allTasks.Exception;
            }

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            Console.WriteLine($"Completed in {TimeSpan.FromMilliseconds(elapsedMs).TotalSeconds} Seconds");
        }



        public static async Task<User> AddGraphUser(GraphServiceClient graphClient, User user)
        {

            try
            {
                return await graphClient.Users.Request().AddAsync(user).ConfigureAwait(false);
            }
            catch (Exception e)
            {
                failedUserList.Add(user);
                throw e;
            }
        }

        public static async Task<User> UpdateGraphUser(GraphServiceClient graphClient, User user)
        {

            try
            {
                return await graphClient.Users[user.Id].Request().UpdateAsync(new User()
                {
                    GivenName = user.GivenName,
                    Surname = user.Surname,
                    DisplayName = user.DisplayName,
                    City = user.City,
                    CompanyName = user.CompanyName
                }).ConfigureAwait(false);
            }
            catch (Exception e)
            {
                failedUserList.Add(user);
                throw e;
            }
        }


        public static async Task BulkCreate(AppSettings config, GraphServiceClient graphClient)
        {
            // Get the users to import
            string appDirectoryPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string dataFilePath = Path.Combine(appDirectoryPath, config.UsersFileName);

            // Verify and notify on file existence
            if (!System.IO.File.Exists(dataFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"File '{dataFilePath}' not found.");
                Console.ResetColor();
                Console.ReadLine();
                return;
            }

            Console.WriteLine("Starting bulk create operation...");

            // Read the data file and convert to object
            UsersModel users = UsersModel.Parse(System.IO.File.ReadAllText(dataFilePath));

            foreach (var user in users.Users)
            {
                user.SetB2CProfile(config.TenantId);

                try
                {
                    // Create the user account in the directory
                    User user1 = await graphClient.Users
                                    .Request()
                                    .AddAsync(user);

                    Console.WriteLine($"User '{user.DisplayName}' successfully created.");
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(ex.Message);
                    Console.ResetColor();
                }
            }
        }

        public static async Task CreateRandomGraphUserBatch(List<UserModel> userList, GraphServiceClient graphClient)
        {

            int maxNoBatchItems = 20;
            List<BatchRequestContent> batches = new List<BatchRequestContent>();
            var batchRequestContent = new BatchRequestContent();
            List<User> createdUsers = new List<User>();

            foreach (var u in userList)
            {
                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, $"https://graph.microsoft.com/v1.0/users")
                {
                    Content = new StringContent(JsonConvert.SerializeObject(u), Encoding.UTF8, "application/json")
                };
                BatchRequestStep requestStep = new BatchRequestStep(userList.IndexOf(u).ToString(), httpRequestMessage, null);
                batchRequestContent.AddBatchRequestStep(requestStep);

                // Max number of 20 request per batch. So we need to send out multiple batches.
                if (userList.IndexOf(u) > 0 && userList.IndexOf(u) % maxNoBatchItems == 0)
                {
                    batches.Add(batchRequestContent);
                    batchRequestContent = new BatchRequestContent();
                }
            }

            if (batchRequestContent.BatchRequestSteps.Count < maxNoBatchItems)
            {
                batches.Add(batchRequestContent);
            }

            if (batches.Count == 0 && batchRequestContent != null) batches.Add(batchRequestContent);
            var tasks = new List<Task<BatchResponseContent>>();


            foreach (BatchRequestContent batch in batches)
            {
                BatchResponseContent response = null;

                try
                {
                    response = await graphClient.Batch.Request().PostAsync(batch).ConfigureAwait(false);
                }
                catch (Microsoft.Graph.ClientException ex)
                {
                    Console.WriteLine(ex.Message);
                }

                Dictionary<string, HttpResponseMessage> responses = await response.GetResponsesAsync();

                //foreach (string key in responses.Keys)
                //{
                //    HttpResponseMessage httpResponse = await response.GetResponseByIdAsync(key);
                //    var responseContent = await httpResponse.Content.ReadAsStringAsync();

                //    JObject userResponse = JObject.Parse(responseContent);

                //    //var user = (User)userResponse["id"];
                //    //Console.WriteLine($"Response code: {responses[key].StatusCode}-{responses[key].ReasonPhrase}-{eventId}");
                //}

            }

        }




        public static List<UserModel> GenerateRandomGraphUsers(int GeneratedUserCount = 1000)
        {
            Random rand = new Random(DateTime.Now.Second);

            RandomName nameGen = new RandomName(rand);
            List<string> Names = nameGen.RandomNames(GeneratedUserCount, 0);
            RandomData rd = new RandomData();

            var userList = new List<UserModel>();

            foreach (var name in Names)
            {
                var user = new UserModel();
                var username = name.Split(' ');
                var email = $"{username[0]}.{username[1]}@{rd.RandomHostName(10)}";
                user.GivenName = username[0];
                user.Surname = username[1];
                user.DisplayName = $"[TEST] {username[0]} {username[1]} (Local account)";
                user.Password = "Pass!w0rd";
                var Identities = new List<ObjectIdentity>();
                Identities.Add(new ObjectIdentity() { IssuerAssignedId = email, SignInType = "emailAddress" });
                user.Identities = Identities;
                ;
                userList.Add(user);
            }

            return userList;
        }

        public static User UpdateTestUserProperties(User u)
        {
            Random rand = new Random(DateTime.Now.Second);

            RandomName nameGen = new RandomName(rand);
            List<string> Names = nameGen.RandomNames(1, 0);
            RandomData rd = new RandomData();
            var name = Names[0];


            var username = name.Split(' ');
            var email = $"{username[0]}.{username[1]}@{rd.RandomHostName(10)}";
            u.GivenName = username[0];
            u.Surname = username[1];
            u.DisplayName = $"[TEST] {username[0]} {username[1]} (Local account)";
            u.City = "Chicago";
            u.CompanyName = "Contoso";
            //user.Password = "Pass!w0rd";
            //var Identities = new List<ObjectIdentity>();
            //Identities.Add(new ObjectIdentity() { IssuerAssignedId = email, SignInType = "emailAddress" });
            //u.Identities = Identities;

            return u;
        }

        public static async Task CreateUserWithCustomAttribute(GraphServiceClient graphClient, string b2cExtensionAppClientId, string tenantId)
        {
            if (string.IsNullOrWhiteSpace(b2cExtensionAppClientId))
            {
                throw new ArgumentException("B2C Extension App ClientId (ApplicationId) is missing in the appsettings.json. Get it from the App Registrations blade in the Azure portal. The app registration has the name 'b2c-extensions-app. Do not modify. Used by AADB2C for storing user data.'.", nameof(b2cExtensionAppClientId));
            }

            // Declare the names of the custom attributes
            const string customAttributeName1 = "FavouriteSeason";
            const string customAttributeName2 = "LovesPets";

            // Get the complete name of the custom attribute (Azure AD extension)
            Helpers.B2cCustomAttributeHelper helper = new Helpers.B2cCustomAttributeHelper(b2cExtensionAppClientId);
            string favouriteSeasonAttributeName = helper.GetCompleteAttributeName(customAttributeName1);
            string lovesPetsAttributeName = helper.GetCompleteAttributeName(customAttributeName2);

            Console.WriteLine($"Create a user with the custom attributes '{customAttributeName1}' (string) and '{customAttributeName2}' (boolean)");

            // Fill custom attributes
            IDictionary<string, object> extensionInstance = new Dictionary<string, object>();
            extensionInstance.Add(favouriteSeasonAttributeName, "summer");
            extensionInstance.Add(lovesPetsAttributeName, true);

            try
            {
                // Create user
                var result = await graphClient.Users
                .Request()
                .AddAsync(new User
                {
                    GivenName = "Casey",
                    Surname = "Jensen",
                    DisplayName = "Casey Jensen",
                    Identities = new List<ObjectIdentity>
                    {
                        new ObjectIdentity()
                        {
                            SignInType = "emailAddress",
                            Issuer = tenantId,
                            IssuerAssignedId = "casey.jensen@example.com"
                        }
                    },
                    PasswordProfile = new PasswordProfile()
                    {
                        Password = Helpers.PasswordHelper.GenerateNewPassword(4, 8, 4)
                    },
                    PasswordPolicies = "DisablePasswordExpiration",
                    AdditionalData = extensionInstance
                });

                string userId = result.Id;

                Console.WriteLine($"Created the new user. Now get the created user with object ID '{userId}'...");

                // Get created user by object ID
                result = await graphClient.Users[userId]
                    .Request()
                    .Select($"id,givenName,surName,displayName,identities,{favouriteSeasonAttributeName},{lovesPetsAttributeName}")
                    .GetAsync();

                if (result != null)
                {
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine($"DisplayName: {result.DisplayName}");
                    Console.WriteLine($"{customAttributeName1}: {result.AdditionalData[favouriteSeasonAttributeName].ToString()}");
                    Console.WriteLine($"{customAttributeName2}: {result.AdditionalData[lovesPetsAttributeName].ToString()}");
                    Console.WriteLine();
                    Console.ResetColor();
                    Console.WriteLine(JsonConvert.SerializeObject(result, Formatting.Indented));
                }
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode == System.Net.HttpStatusCode.BadRequest)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Have you created the custom attributes '{customAttributeName1}' (string) and '{customAttributeName2}' (boolean) in your tenant?");
                    Console.WriteLine();
                    Console.WriteLine(ex.Message);
                    Console.ResetColor();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }

        public static async Task ForEach<T>(ICollection<T> source, Func<T, Task> body, CancellationToken token)
        {
            // create the list of tasks we will be running
            var tasks = new List<Task>(source.Count);
            try
            {
                // and add them all at once.
                tasks.AddRange(source.Select(s => Task.Run(() => body(s), token)));

                // execute it all with a delay to throw.
                for (; ; )
                {
                    // very short delay
                    var delay = Task.Delay(1, token);

                    // and all our tasks
                    await Task.WhenAny(Task.WhenAll(tasks), delay).ConfigureAwait(false);
                    if (tasks.All(t => t.IsCompleted))
                    {
                        break;
                    }

                    //
                    // ... use a spinner or something
                }
                await Task.WhenAll(tasks.ToArray()).ConfigureAwait(false);

                // throw if we are done here.
                token.ThrowIfCancellationRequested();
            }
            finally
            {
                // find the error(s) that might have happened.
                var errors = tasks.Where(tt => tt.IsFaulted).Select(tu => tu.Exception).ToList();

                // we are back in our own thread
                if (errors.Count > 0)
                {
                    throw new AggregateException(errors);
                }
            }
        }
    }

    public class DistinctUserComparer : IEqualityComparer<UserModel>
    {

        public bool Equals(UserModel x, UserModel y)
        {
            return x.GivenName == y.GivenName &&
                x.Surname == y.Surname;
        }

        public int GetHashCode(UserModel obj)
        {
            return obj.GivenName.GetHashCode() ^
                obj.Surname.GetHashCode();
        }
    }
}
