using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailMonitoringService
{
    public class Logger
    {
        public static void WriteEmailLog(string message)
        {
            using (StreamWriter writer = new StreamWriter(GlobalsVariables.logEmailPath, true, Encoding.UTF8))
            {
                writer.WriteLine($"{DateTime.Now} :{message}");
            }
        }
        public static void WriteSystemLog(string message)
        {
            using (StreamWriter writer = new StreamWriter(GlobalsVariables.logSystemPath, true, Encoding.UTF8))
            {
                writer.WriteLine($"{DateTime.Now} :{message}");
            }
        }

        public static void ReadLastSystemLog(string word)
        {
            var lines = File.ReadAllLines(GlobalsVariables.logSystemPath);

            // Filtracja linii zawierających jakieś słowo 

            var lastLine = lines
                .Where(line => line.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0).LastOrDefault();
            if (lastLine != null)
            {
                WriteSystemLog($"Ostatnie wyłaczenie się programu: {lastLine}");
            }
        }

        public static void CleanOldLogs()
        {
            // Odczytanie wszystkich linii z pliku logów
            string[] logLines = File.ReadAllLines(GlobalsVariables.logEmailPath);

            // Obliczanie daty granicznej (1 dzień temu)
            DateTime thresholdDate = DateTime.Now.AddDays(-1);

            // Filtracja linii, które mają datę nowszą niż 1 dzień
            var filteredLogLines = logLines.Where(line =>
            {
                // Wyodrębnienie daty i godziny z początku linii
                string dateString = line.Substring(0, 10); // `dd.MM.yyyy`
                string timeString = line.Substring(11, 8); // `HH:mm:ss`
                string dateTimeString = $"{dateString} {timeString}";

                if (DateTime.TryParseExact(dateTimeString, GlobalsVariables.dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime logDate))
                {
                    return logDate >= thresholdDate;
                }
                return false;
            }).ToArray(); // Konwersja do tablicy

            // Zapisanie przefiltrowanych linii do pliku logów
            File.WriteAllLines(GlobalsVariables.logEmailPath, filteredLogLines);
        }
        public static void CompareDate()
        {

            // Odczytanie wszystkich linii z pliku logów
            string[] logLines = File.ReadAllLines(GlobalsVariables.logSystemPath);
            if (logLines.Length < 2)
            {
                Console.WriteLine("Tablica nie zawiera wystarczającej liczby elementów do porównania.");
                SMTP.SendEmail("Problem z plikiem SystemLog", "Brak przynajmniej 2 wpisów w pliku");
            }
            // Zapisanie przefiltrowanych linii do pliku logów do ustawienia w zalożeności ile logów ustawiło się na starcie
            string ostatniLine = logLines[logLines.Length - 1];

            DateTime ostatniDateTime = DateTime.ParseExact(ostatniLine.Substring(0, 19), GlobalsVariables.dateFormat, CultureInfo.InvariantCulture);

            // Obliczanie różnicy czasu
            TimeSpan timeDifference = DateTime.Now - ostatniDateTime;

            WriteSystemLog($"Program był wyłączony przez {timeDifference.Days} dni, {timeDifference.Hours} godzin, {timeDifference.Minutes} minut, {timeDifference.Seconds} sekund");

        }

    }
}



