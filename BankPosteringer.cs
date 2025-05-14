using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVM_Kasserer
{
    class BankPosteringer
    {
        private string bankPosteringerFilepath = @"C:\Users\kaspe\Downloads\eksport.csv";

        string matchesFilePath = @"D:\Dropbox\HVM - Kasserer\Program-kode\HVM Kasserer\cprToAddressMatches.xlsx";

        List<CprToSenderMatch> cprToSenderMatches;

        public BankPosteringer()
        {
            cprToSenderMatches = LoadMatchesFromFile();
        }

        public void HandleBankPosteringer()
        {
            // Get all of the transactions parsed to an object we can use
            var transactionData = new List<BankPostering>();
            using (var reader = new StreamReader(bankPosteringerFilepath))
            {
                // Skip the header row
                reader.ReadLine();

                CultureInfo cultureInfo = CultureInfo.GetCultureInfo("da-DK");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (line == null) break;

                    string[] values = line.Split(';');

                    transactionData.Add(new BankPostering
                    {
                        Date = DateTime.Parse(values[0], cultureInfo),
                        Message = values[1],
                        Amount = decimal.Parse(values[2], cultureInfo),
                        Address = values[6]
                    });
                }
            }

            // Only use transactions with positive amounts
            transactionData = transactionData.Where(t => t.Amount > 0).ToList();

            // traverse transactionData and find all destinct addresses, that is a multiline address
            var addresses = transactionData
                .Select(t => t.Address)
                .Distinct()
                .Where(a => a.Contains(", "))
                .ToList();

            // For each address, find the corresponding CPR number, if none are found ask the user for the CPR number and add the pair to the file
            foreach (var address in addresses)
            {
                var cprMatch = cprToSenderMatches.FirstOrDefault(m => m.SenderAddress.Equals(address, StringComparison.OrdinalIgnoreCase));
                if (cprMatch == null)
                {
                    Console.WriteLine($"No CPR match found for address: {address}");
                    Console.Write("Do you want to add a match entry for that address? y/n: ");
                    string? addEntry = Console.ReadLine();
                    if (string.IsNullOrEmpty(addEntry) || addEntry?.ToLower() != "y")
                    {
                        continue; // Skip to the next address
                    }

                    // if two deposits are found on the same day, with the same address but different messages, it indicates we should add two entries as two people can live on the same address
                    var depositsGroupedByDate = transactionData
                        .Where(t => t.Address.Equals(address, StringComparison.OrdinalIgnoreCase))
                        .GroupBy(t => t.Date.Date)
                        .ToList();

                    // If Any of the entries in the variable deposists has multiple entries, we should ask for each entry and present both address and message
                    var depositsWithMutlipleEntriesOnSameDay = depositsGroupedByDate.Where(x => x.Count() > 1).ToList();

                    if (depositsWithMutlipleEntriesOnSameDay.Any(x => x.Count() > 1))
                    {
                        var firstDeposit = depositsGroupedByDate.First();
                        foreach (var deposit in firstDeposit)
                        {
                            string? cpr = null;
                            while (string.IsNullOrWhiteSpace(cpr))
                            {
                                Console.Write($"Please enter the CPR number for this message: \"{deposit.Message}\": ");
                                cpr = Console.ReadLine();

                                if (string.IsNullOrWhiteSpace(cpr))
                                {
                                    Console.WriteLine("Input cannot be empty. Please try again.");
                                }
                            }

                            using var workbook = new XLWorkbook(matchesFilePath);
                            var worksheet = workbook.Worksheet(1);

                            var rowNumber = worksheet.RowsUsed().Last().RowNumber();
                            var newRow = worksheet.Row(rowNumber + 1);

                            newRow.Cell(1).Value = cpr;
                            newRow.Cell(2).Value = address;
                            newRow.Cell(3).Value = deposit.Message;

                            // Save the workbook
                            workbook.Save();
                        }
                    }
                    else
                    {
                        string? cpr = null;
                        while (string.IsNullOrWhiteSpace(cpr))
                        {
                            Console.Write($"Please enter the CPR number: ");
                            cpr = Console.ReadLine();

                            if (string.IsNullOrWhiteSpace(cpr))
                            {
                                Console.WriteLine("Input cannot be empty. Please try again.");
                            }
                        }

                        using var workbook = new XLWorkbook(matchesFilePath);
                        var worksheet = workbook.Worksheet(1);

                        var rowNumber = worksheet.RowsUsed().Last().RowNumber();
                        var newRow = worksheet.Row(rowNumber + 1);

                        newRow.Cell(1).Value = cpr;
                        newRow.Cell(2).Value = address;

                        // Save the workbook
                        workbook.Save();
                    }
                }
            }
        }

        private List<CprToSenderMatch> LoadMatchesFromFile()
        {
            if (!File.Exists(matchesFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Matches file not found at: {matchesFilePath}");
                Console.ForegroundColor = ConsoleColor.Gray;
                return new List<CprToSenderMatch>();
            }

            var matches = new List<CprToSenderMatch>();

            using (var workbook = new XLWorkbook(matchesFilePath))
            {
                var worksheet = workbook.Worksheet(1); // Assuming the matches are in the first sheet

                foreach (var row in worksheet.RowsUsed().Skip(1)) // Skip the header row
                {
                    string cpr = row.Cell(1).GetValue<string>().Trim(); // Column A
                    string address = row.Cell(2).GetValue<string>().Trim(); // Column B
                    string keywords = row.Cell(3).GetValue<string>().Trim(); // Column C

                    if (!string.IsNullOrWhiteSpace(cpr) && !string.IsNullOrWhiteSpace(address))
                    {
                        matches.Add(new CprToSenderMatch
                        {
                            CPR = cpr,
                            SenderAddress = address,
                            Keywords = !string.IsNullOrWhiteSpace(keywords)
                                ? keywords.Split(',').Select(k => k.Trim()).ToList()
                                : new List<string>()
                        });
                    }
                }
            }

            return matches;
        }

        class CprToSenderMatch
        {
            public required string CPR { get; set; }
            public required string SenderAddress { get; set; }
            public List<string>? Keywords { get; set; }
        }

        class BankPostering
        {
            public DateTime Date { get; set; }
            public required string Message { get; set; }
            public decimal Amount { get; set; }
            public required string Address { get; set; }
        }
    }
}
