using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace EmailMonitoringService
{
    static class SMTP
    {
        public static void SendEmail(string email, string subject, string body)

        {
            try
            {
                // Konfiguracja klienta SMTP
                using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587))
                {
                    client.Credentials = new NetworkCredential(GlobalsVariables.email, GlobalsVariables.appPassword);
                    client.EnableSsl = true;


                    // Konfiguracja wiadomości e-mail
                    MailMessage mailMessage = new MailMessage
                    {
                        From = new MailAddress(GlobalsVariables.email),
                        Subject = subject,
                        Body = body

                    };
                    mailMessage.To.Add(email);

                    // Wysyłka e-maila
                    client.Send(mailMessage);
                    Console.WriteLine("E-mail wysłany.");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteSystemLog($"Wystąpił błąd podczas wysyłania e-maila: {ex.Message}");
            }

        }
    }
}