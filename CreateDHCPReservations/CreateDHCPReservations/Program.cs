using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using CsvHelper;

namespace CreateDHCPReservations
{
    public class Program
    {
        private static string logsFolder = "C:\\temp\\dhcplogs";
        private static string reservationsOutputFile = "c:\\temp\\DHCPReservations.csv";
        private static string scopeId = "10.10.0.0";

        static void Main(string[] args)
        {
            try
            {
                if(args.Length!= 1)
                    throw new ArgumentException("program requires read or create param");

                var step = args[0];

                switch (step.ToLower())
                {
                    case "read":
                        CreateReservationsOutputFile();
                        break;
                    case "create":
                        CreateReservationsInDHCPFromCsv();
                        break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static void CreateReservationsInDHCPFromCsv()
        {
            var reservations = ReadReservationsCsv();

            foreach (var reservation in reservations.Skip(1))
            {
                var command = $"netsh dhcp server scope {reservation.ScopeId} add reservedip {reservation.IpAddress} {reservation.MacAddress} {reservation.MachineName} \"{reservation.MachineName}\"";

                var procStartInfo = new ProcessStartInfo("cmd", "/c " + command);
                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                procStartInfo.CreateNoWindow = true;

                using (var proc = new Process())
                {
                    proc.StartInfo = procStartInfo;
                    proc.Start();
                    var result = proc.StandardOutput.ReadToEnd();
                    Console.WriteLine(result);
                }
            }
        }

        private static void CreateReservationsOutputFile()
        {
            var logs = ReadLogFiles();

            var deleted = logs.Where(x => x.Description == "Deleted").ToList();

            var result = new Dictionary<string, DhcpReservation>();

            foreach (var deletedLogRecords in deleted)
            {
                if (result.ContainsKey(deletedLogRecords.IpAddress)) continue;

                var output = new DhcpReservation
                {
                    IpAddress = deletedLogRecords.IpAddress,
                    MacAddress = deletedLogRecords.MACAddress,
                    MachineName = FindMachineName(deletedLogRecords, logs),
                    ScopeId = scopeId
                };

                result.Add(output.IpAddress, output);
            }

            WriteOutput(result.Select(x => x.Value).ToList());
        }

        private static List<DhcpLog> ReadLogFiles()
        {
            var files = Directory.GetFiles(logsFolder, "DhcpSrvLog-*.log");

            var result = new List<DhcpLog>();

            foreach (var file in files)
            {
                result.AddRange(ReadLogFile(file));
            }

            return result;
        }

        private static IEnumerable<DhcpLog> ReadLogFile(string logFile)
        {
            using (var reader = new StreamReader(logFile))
            using (var csv = new CsvReader(reader))
            {
                csv.Configuration.HasHeaderRecord = false;
                return csv.GetRecords<DhcpLog>().ToList();
            }
        }

        private static IEnumerable<DhcpReservation> ReadReservationsCsv()
        {
            using (var reader = new StreamReader(reservationsOutputFile))
            using (var csv = new CsvReader(reader))
            {
                csv.Configuration.HasHeaderRecord = false;
                return csv.GetRecords<DhcpReservation>().ToList();
            }
        }

        private static string FindMachineName(DhcpLog deleteRecord, IEnumerable<DhcpLog> data)
        {
            foreach (var dhcpRecord in data)
            {
                if (dhcpRecord.IpAddress == deleteRecord.IpAddress)
                {
                    if (!string.IsNullOrEmpty(dhcpRecord.HostName))
                        return dhcpRecord.HostName;
                }
            }

            return string.Empty;
        }

        private static void WriteOutput(IEnumerable<DhcpReservation> output)
        {
            using (var writer = new StreamWriter(reservationsOutputFile))
            using (var csv = new CsvWriter(writer))
            {
                csv.WriteRecords(output);
            }
        }
    }
}
