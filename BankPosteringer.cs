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

        string matchesFilePath = @"C:\Dropbox\HVM - Kasserer\Program-kode\HVM Kasserer\cprToAddressMatches.xlsx";

        List<CprToSenderMatch> cprToSenderMatches;

        public BankPosteringer()
        {
            cprToSenderMatches = LoadMatchesFromFile();
        }

        public void HandleBankPosteringer()
        {
            // Get all of the transactions parsed to an object we can use
            List<BankPostering> transactionData = ExtractTransactionsDataFromCSV();

            // Sort out transctions that are not from members
            transactionData = transactionData.Where(t => 
                !t.Address.Contains("Vipps", StringComparison.OrdinalIgnoreCase) &&
                !t.Address.Contains("Begravelseshjælp", StringComparison.OrdinalIgnoreCase) &&
                !t.Address.Contains("Korskærvej 25, 7000", StringComparison.OrdinalIgnoreCase) &&
                !t.Address.Contains("Kirkegade 15, 8722  Hedensted", StringComparison.OrdinalIgnoreCase)
            ).ToList();

            // Only use transactions with positive amounts
            transactionData = transactionData.Where(t => t.Amount > 0).ToList();

            // Only transactions that are not kontingent
            transactionData = transactionData.Where(t => !t.Message.Contains("kontingent", StringComparison.OrdinalIgnoreCase) && !t.Message.Contains("kont", StringComparison.OrdinalIgnoreCase)).ToList();

            AddAddressAndCPRPairsToFile(transactionData);

            // summerize the transactions by month
            var transactionsByMonth = transactionData
                .GroupBy(t => new { t.Date.Year, t.Date.Month })
                .Select(g => new
                {
                    g.Key.Year,
                    g.Key.Month,
                    TotalAmount = g.Sum(t => t.Amount)
                })
                .ToList();

            // Print the summary
            Console.WriteLine("Monthly Summary:");
            foreach (var month in transactionsByMonth)
            {
                Console.WriteLine($"Year: {month.Year}, Month: {month.Month}, Total Amount: {month.TotalAmount}");
            }

            // For every month, group the transactions by address and sum the amounts for each address
            var transactionsByAddress = transactionData
                .GroupBy(t => new { t.Date.Year, t.Date.Month, t.Address })
                .Select(g => new
                {
                    g.Key.Year,
                    g.Key.Month,
                    g.Key.Address,
                    TotalAmount = g.Sum(t => t.Amount)
                })
                .ToList();

            // Print the summary
            Console.WriteLine("\nMonthly Summary by Address:");
            foreach (var month in transactionsByAddress)
            {
                Console.WriteLine($"Year: {month.Year}, Month: {month.Month}, Address: {month.Address}, Total Amount: {month.TotalAmount}");
            }

        }

        private void AddAddressAndCPRPairsToFile(List<BankPostering> transactionData)
        {
            // traverse transactionData and find all destinct addresses, that is a multiline address
            List<string> addresses = transactionData
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

        private static bool StartsWithDate(string line)
        {
            if (string.IsNullOrWhiteSpace(line) || line.Length < 10)
                return false;

            string first10 = line.Substring(0, 10);
            return DateTime.TryParseExact(first10, "dd-MM-yyyy",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out _);
        }

        private List<BankPostering> ExtractTransactionsDataFromCSV()
        {
            var transactionData = new List<BankPostering>();
            using (var reader = new StreamReader(bankPosteringerFilepath))
            {

                string currentLine = null;
                string nextLine = null;

                CultureInfo cultureInfo = CultureInfo.GetCultureInfo("da-DK");

                // Pseudocode:
                // - Read a line from the CSV.
                // - If the next line does not start with a date (format "dd-MM-yyyy"), append it to the current line.
                // - Continue until a line starting with a date is found or end of file.
                // - Then process the combined line as usual.

                while ((currentLine = currentLine ?? reader.ReadLine()) != null)
                {
                    nextLine = reader.ReadLine();

                    while (nextLine != null && !StartsWithDate(nextLine))
                    {
                        currentLine += " " + nextLine; // Or use newline or another delimiter
                        nextLine = reader.ReadLine();
                    }

                    string[] values = currentLine.Split(';');

                    transactionData.Add(new BankPostering
                    {
                        Date = DateTime.Parse(values[0], cultureInfo),
                        Message = values[1],
                        Amount = decimal.Parse(values[2], cultureInfo),
                        Address = values[7]
                    });

                    currentLine = nextLine;
                }
            }

            return transactionData;
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
