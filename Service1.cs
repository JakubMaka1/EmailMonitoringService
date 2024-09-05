using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MailKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.CompilerServices;

namespace EmailMonitoringService
{
    public static class GlobalsVariables
    {
        public static string email = Environment.GetEnvironmentVariable("EMAIL_ADDRESS");
        public static string password = Environment.GetEnvironmentVariable("EMAIL_PASSWORD");
        public static string appPassword = Environment.GetEnvironmentVariable("KodAplikacji");
        public static string logEmailPath = "C:/Temp/EmailLog.txt";
        public static string logSystemPath = "C:/Temp/SystemLog.txt";
        public static string filePolicyPath = "C:/Temp/Policy.txt";
        public static string dateFormat = "dd.MM.yyyy HH:mm:ss";
    }

    [RunInstaller(true)]
    public partial class Service1 : ServiceBase
    {
        private CancellationTokenSource _cancellationTokenSource;
        private Task _monitoringTask;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            ProgramRuntime programRuntime = new ProgramRuntime();
            
            _cancellationTokenSource = new CancellationTokenSource();
            _monitoringTask = Task.Run(() => StartMonitoring(_cancellationTokenSource.Token));
        }

        protected override void OnStop()
        {
            ProgramRuntime programRuntime = new ProgramRuntime();
            programRuntime.StopAndDisplayRuntime();
            _cancellationTokenSource.Cancel();
            SMTP.SendEmail("Error", "Wyłaczenie programu");
            Logger.WriteSystemLog("Wyłaczenie programu");
        }

        private async Task StartMonitoring(CancellationToken cancellationToken)
        {            
            
            Logger.CompareDate();

            if (GlobalsVariables.email == null || GlobalsVariables.appPassword == null)
            {
                Logger.WriteEmailLog("Email or password environment variable is not set.");
                SMTP.SendEmail("Error", "Email or password environment variable is not set.");
            }
        Start:
            using (var client = new ImapClient())
            {                
                try
                {
                    await client.ConnectAsync("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);
                    await client.AuthenticateAsync(GlobalsVariables.email, GlobalsVariables.appPassword);
                    var inbox = client.Inbox;
                    await inbox.OpenAsync(FolderAccess.ReadWrite);

                    Logger.WriteSystemLog("Monitoring folderu INBOX...");

                    while (!cancellationToken.IsCancellationRequested)
                    {
                        await inbox.CheckAsync(cancellationToken);
                        var uids = await inbox.SearchAsync(SearchQuery.NotSeen);

                        foreach (var uid in uids)
                        {
                            Logger.CleanOldLogs();

                            var message = await inbox.GetMessageAsync(uid);
                            Logger.WriteEmailLog($"Odebrano wiadomość: {message.Subject}");

                            if (message.Subject.IndexOf("disable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                await inbox.AddFlagsAsync(uid, MessageFlags.Seen, true);
                                string jsonData = "{ \"status\": \"disable\" }";
                                await SendPutRequest(jsonData);
                                await SendGETRequest();
                                
                            }
                            
                            else if (message.Subject.IndexOf("enable", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                await inbox.AddFlagsAsync(uid, MessageFlags.Seen, true);
                                string jsonData = "{ \"status\": \"enable\" }";
                                await SendPutRequest(jsonData);
                                await SendGETRequest();
                                
                            }

                            await inbox.AddFlagsAsync(uid, MessageFlags.Seen, true);
                        }

                        await Task.Delay(10000, cancellationToken);
                    }

                    await client.DisconnectAsync(true);
                }
                catch (OperationCanceledException)
                {
                    ProgramRuntime programRuntime = new ProgramRuntime();
                    Console.WriteLine("Operacja anulowana.");
                    Logger.WriteSystemLog("Operacja anulowana.");
                    programRuntime.ErrorStopAndDisplayRuntime("Operacja anulowana");
                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Wystąpił błąd: {ex.Message}");
                    Logger.WriteSystemLog($"Wystąpił błąd catch ex startmonitoring : {ex.Message}");
                    SMTP.SendEmail("Error z połaczeniem z FG", $"Wystąpił błąd: {ex.Message}");
                    //programRuntime.ErrorStopAndDisplayRuntime($"Exceprion ex {ex.Message}");
                    goto Start;
                }
            }
        }

        private async Task SendPutRequest(string jsonData)
        {           
                using (var httpClient = new HttpClient())
                {
                    string[] policyLines = File.ReadAllLines(GlobalsVariables.filePolicyPath);

                    foreach (var policy in policyLines)
                    {
                        var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                        var response = await httpClient.PutAsync($"http://10.10.10.4/api/v2/cmdb/firewall/policy/{policy}/?access_token=gtjmQkf4s1kQwQ9yd3nhmfxx39ykb9", content);

                        if (response.IsSuccessStatusCode)
                        {
                            Logger.WriteSystemLog("Operacja aktualizacji statusu udana.");
                        }
                        else
                        {
                            Logger.WriteSystemLog($"Wystąpił błąd else sendput: {response.StatusCode}");
                            SMTP.SendEmail("Error_PUT", $"Wystąpił błąd: {response.StatusCode}");
                      
                        }
                    }
                }            
        }

        private async Task SendGETRequest()
        {
            using (var httpClient = new HttpClient())
            {
                string[] policyLines = File.ReadAllLines(GlobalsVariables.filePolicyPath);

                foreach (var policy in policyLines)
                {
                    HttpResponseMessage response = await httpClient.GetAsync($"http://10.10.10.4/api/v2/cmdb/firewall/policy/{policy}?access_token=gtjmQkf4s1kQwQ9yd3nhmfxx39ykb9");
                    string json = await response.Content.ReadAsStringAsync();
                    JObject jsonObject = JObject.Parse(json);

                    foreach (var result in jsonObject["results"])
                    {
                        if ((string)result["policyid"] == policy)
                        {
                            Logger.WriteEmailLog($"Nazwa policy: {result["name"]}");
                            Logger.WriteEmailLog($"Aktualny status: {result["status"]}");
                            Logger.WriteSystemLog($"Nazwa policy: {result["name"]}");
                            Logger.WriteSystemLog($"Aktualny status: {result["status"]}");
                        }
                    }
                }
            }
        }
    }
}
