using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;

class Program
{
    private static string mobilePayFilepath = @"D:\Dropbox\HVM - Kasserer\Indsamlinger\2025 Indsamlinger\transactions-report-1_1_2024-12_29_2024.csv";
    static List<string> exclusions = LoadExclusionsFromFile();

    private static List<string> LoadExclusionsFromFile()
    {
        string solutionDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string exclusionsFilePath = Path.Combine(solutionDirectory, "mobilePayExclusions.txt");

        if (!File.Exists(exclusionsFilePath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Exclusions file not found at: {exclusionsFilePath}");
            Console.ForegroundColor = ConsoleColor.Gray;
            return new List<string>();
        }

        return File.ReadAllLines(exclusionsFilePath)
                   .Where(line => !string.IsNullOrWhiteSpace(line))
                   .Select(line => line.Trim())
                   .ToList();
    }

    private static string excelFilepath = @"D:\Dropbox\HVM - Kasserer\Indsamlinger\2025 Indsamlinger\2025 Mobilepay - Copy.xlsx";

    static void Main()
    {
        SummarizeMobilePayTransactions();
    }

    static void SummarizeMobilePayTransactions()
    {
        var transactions = GetTransactions();

        // Step 1: Identify TransactionIDs with exclusion keywords
        var excludedTransactionIDs = transactions
            .Where(t => t.Message != null &&
                        exclusions.Any(keyword => NormalizeString(t.Message).Contains(NormalizeString(keyword))))
            .Select(t => t.TransactionID)
            .Distinct()
            .ToHashSet();

        // Mark transactions as "excluded" or "regular"
        foreach (var transaction in transactions)
        {
            transaction.IsExcluded = excludedTransactionIDs.Contains(transaction.TransactionID);
            transaction.Name = transaction.Name.Replace("  ", " "); // Replace multiple spaces with a single one
        }

        // Summarize by day
        var dailySummary = transactions
            .GroupBy(t => t.Date.Date)
            .Select(group =>
            {
                var regularTransactions = group.Where(t => !t.IsExcluded).ToList();
                var excludedTransactions = group.Where(t => t.IsExcluded).ToList();

                return new
                {
                    Date = group.Key,
                    CombinedTotal = group.Sum(t => t.Amount),

                    RegularDonation = regularTransactions.Where(x => x.Type != "Gebyr").Sum(x => x.Amount),
                    RegularTotal = regularTransactions.Sum(t => t.Amount),
                    RegularGebyr = regularTransactions.Where(x => x.Type == "Gebyr").Sum(x => x.Amount) * -1,
                    RegularMessages = string.Join(", ", regularTransactions.Where(t => t.Message != null && t.Message.Length > 0).Select(t => t.Message).Distinct()),

                    ExcludedDonation = excludedTransactions.Where(x => x.Type != "Gebyr").Sum(x => x.Amount),
                    ExcludedTotal = excludedTransactions.Sum(t => t.Amount),
                    ExcludedGebyr = excludedTransactions.Where(x => x.Type == "Gebyr").Sum(x => x.Amount) * -1,
                    ExcludedMessages = string.Join(", ", excludedTransactions.Select(t => t.Message).Distinct())
                };
            })
            .OrderBy(x => x.Date)
            .ToList();


        // Output daily summary
        foreach (var day in dailySummary)
        {
            Console.WriteLine($"Date: {day.Date}");
            if (day.RegularTotal > 0 && day.ExcludedTotal > 0)
            {
                Console.WriteLine($"  Combined Total: {day.CombinedTotal}");
                Console.WriteLine($"");
            }

            if (day.RegularTotal > 0)
            {
                Console.WriteLine($"  Regular Donation: {day.RegularDonation}");
                Console.WriteLine($"  Regular Gebyr: {day.RegularGebyr}");
                Console.WriteLine($"  Regular Total: {day.RegularTotal}");
                Console.WriteLine($"  Regular Messages: {day.RegularMessages}");
                Console.WriteLine($"");
            }

            if (day.ExcludedTotal > 0)
            {
                Console.WriteLine($"  Excluded Donation: {day.ExcludedDonation}");
                Console.WriteLine($"  Excluded Gebyr: {day.ExcludedGebyr}");
                Console.WriteLine($"  Excluded Total: {day.ExcludedTotal}");
                Console.WriteLine($"  Excluded Messages: {day.ExcludedMessages}");
                Console.WriteLine($"");
            }
            Console.WriteLine($"");
        }

        // Summarize by month for each person (excluding marked transactions)
        var monthlySummary = transactions
            .Where(t => !t.IsExcluded && t.Type != "Gebyr")
            .GroupBy(t => new { t.Date.Year, t.Date.Month, t.Name, t.Phone })
            .Select(group => new MonthlySummary(
                    RearrangeName(group.Key.Name),
                    group.Key.Phone,
                    group.Key.Year,
                    group.Key.Month,
                    group.Sum(t => t.Amount)
                )
            )
            .OrderBy(x => x.Month)
            .ToList();

        var months = monthlySummary.GroupBy(x => x.Month).ToList();

        foreach (var month in months)
        {
            var monthString = GetMonthAsString(month.Key);
            if (monthString == null)
                throw new Exception();

            // Output monthly summary
            foreach (var entry in month.OrderBy(x => x.Person))
            {
                Console.WriteLine($"{entry.Year}-{entry.Month:00}: {entry.Person}-{entry.Number} - {entry.Total}");

                UpdateExcelFile("Mobilepay", entry.Person, entry.Number, monthString, entry.Total);
            }

            Console.WriteLine($"");
        }

        var excludedTransactions = transactions.Where(t => t.IsExcluded && t.Type != "Gebyr");
        var groupedByDay = excludedTransactions.GroupBy(t => t.Date.Date);
        foreach (var date in groupedByDay)
        {
            var todaysTotal = date.Select(t => t.Amount).Sum();
            string dateString = date.Select(x => x.Date).First().ToString("yyyy-MM-dd");
            string? message = date.Where(x => x.Message != null).Select(x => x.Message).First();
            UpdateExcelForExcluded("Mobilepay", todaysTotal, dateString, message);
        }
    }

    private static string? GetMonthAsString(int monthNumber)
    {
        if (monthNumber == 1) return "Jan";
        if (monthNumber == 2) return "Feb";
        if (monthNumber == 3) return "Marts";
        if (monthNumber == 4) return "April";
        if (monthNumber == 5) return "Maj";
        if (monthNumber == 6) return "Juni";
        if (monthNumber == 7) return "Juli";
        if (monthNumber == 8) return "August";
        if (monthNumber == 9) return "Sept";
        if (monthNumber == 10) return "Okt";
        if (monthNumber == 11) return "Nov";
        if (monthNumber == 12) return "Dec";

        return null;
    }

    private static List<Transaction> GetTransactions()
    {
        var transactionData = new List<Transaction>();

        using (var reader = new StreamReader(mobilePayFilepath))
        {
            // Skip the header row
            reader.ReadLine();

            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("da-DK");

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (line == null) break;

                string[] values = line.Split(';');

                transactionData.Add(new Transaction
                {
                    Date = DateTime.Parse(values[10], cultureInfo),
                    Name = values[15],
                    Phone = ExtractLast4Digits(values[16]),
                    Amount = decimal.Parse(values[6], cultureInfo),
                    Message = values[11].Replace('\"', ' ').Replace('\'', ' ').Trim(),
                    Type = values[5],
                    TransactionID = values[14],
                });
            }
        }

        return transactionData;
    }

    static string NormalizeString(string input)
    {
        return Regex.Replace(input.ToLower(), @"\s+", " ").Trim();
    }

    static string RearrangeName(string fullName)
    {
        var nameParts = fullName.Split(' ');

        if (nameParts.Length > 1)
        {
            // Efternavn er den sidste del, fornavn og eventuelle mellemnavne er resten
            var lastName = nameParts.Last();
            var firstNameAndMiddleNames = string.Join(' ', nameParts.Take(nameParts.Length - 1));

            return $"{lastName}, {firstNameAndMiddleNames}";
        }

        return fullName; // Returner navnet uændret, hvis det kun består af et enkelt ord
    }

    static void UpdateExcelFile(string sheetName, string name, string last4PhoneDigits, string month, decimal value)
    {
        // Open the workbook
        using var workbook = new XLWorkbook(excelFilepath);
        var worksheet = workbook.Worksheet(sheetName);

        // Define column indices
        int nameColumnIndex = GetColumnIndex(worksheet, "Fornavne");
        int phoneColumnIndex = GetColumnIndex(worksheet, "Mobil nummer");

        if (nameColumnIndex == -1 || phoneColumnIndex == -1)
        {
            Console.WriteLine("Required columns ('Fornavne' or 'Mobil nummer') not found.");
            return;
        }

        // Find all rows matching the phone number
        var matchingRowsByPhone = worksheet.RowsUsed()
            .Where(row =>
            {
                string rawPhoneNumber = row.Cell(phoneColumnIndex).GetValue<string>().Trim();
                string cleanedPhoneNumber = ExtractLast4Digits(rawPhoneNumber);
                return cleanedPhoneNumber == last4PhoneDigits;
            })
            .ToList();

        IXLRow? personRow = null;

        if (matchingRowsByPhone.Count == 1)
        {
            // Exact match by phone number
            personRow = matchingRowsByPhone[0];
        }
        else
        {
            // Fallback to matching by name
            var matchingRowsByName = worksheet.RowsUsed()
                .Where(row =>
                {
                    string personName = row.Cell(nameColumnIndex).GetValue<string>().Trim();
                    return personName.Equals(name, StringComparison.OrdinalIgnoreCase);
                })
                .ToList();

            if (matchingRowsByName.Count == 1)
            {
                personRow = matchingRowsByName[0];
                string existingPhoneNumber = ExtractLast4Digits(personRow.Cell(phoneColumnIndex).GetValue<string>().Trim());

                if (string.IsNullOrEmpty(existingPhoneNumber))
                {
                    // Assign the phone number if it doesn't already exist
                    personRow.Cell(phoneColumnIndex).Value = last4PhoneDigits;
                }
                else if(existingPhoneNumber != last4PhoneDigits)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ambiguity detected: Phonenumber not found, but antother person with same name was found. Phone number: {last4PhoneDigits}. Name {name}");
                    Console.ForegroundColor = ConsoleColor.Gray;
                    return;
                }
            }
            else if (matchingRowsByName.Count > 1)
            {
                if (!matchingRowsByPhone.Any())
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ambiguity detected: multiple people with the same name and no matching phone number: {last4PhoneDigits}.");
                    Console.ForegroundColor = ConsoleColor.Gray;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ambiguity detected: multiple people with the same name and the same phone number: {last4PhoneDigits}.");
                    Console.ForegroundColor = ConsoleColor.Gray;
                }

                return;
            }
        }

        if (personRow == null)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"No matching person found. Looking for phone {last4PhoneDigits} and name {name}");
            Console.ForegroundColor = ConsoleColor.Gray;
            return;
        }

        // Find the column for the month
        int monthColumnIndex = GetColumnIndex(worksheet, month);

        if (monthColumnIndex == -1)
        {
            Console.WriteLine($"Month '{month}' not found.");
            return;
        }

        // Update the cell value
        int personRowNumber = personRow.RowNumber();

        worksheet.Cell(personRowNumber, monthColumnIndex).Value = value;

        // Save the workbook
        workbook.Save();

        Console.WriteLine("Excel file updated successfully.");
    }

    static int GetColumnIndex(IXLWorksheet worksheet, string columnHeader)
    {
        var headerRow = worksheet.Row(1); // Assuming headers are in the first row
        var headerCell = headerRow.CellsUsed()
            .FirstOrDefault(cell => cell.GetValue<string>().Trim().Equals(columnHeader, StringComparison.OrdinalIgnoreCase));

        return headerCell?.Address.ColumnNumber ?? -1;
    }

    static string ExtractLast4Digits(string phoneNumber)
    {
        // Remove spaces and non-numeric characters
        var digits = new string(phoneNumber.Where(char.IsDigit).ToArray());

        // Return the last 4 digits if available
        return digits.Length >= 4 ? digits[^4..] : string.Empty;
    }

    // Method to insert excluded transactions above the "I alt" line
    static void UpdateExcelForExcluded(string sheetName, decimal excludedAmount, string date, string? messages)
    {
        using var workbook = new XLWorkbook(excelFilepath);
        var worksheet = workbook.Worksheet(sheetName);

        // Find the row that says "I alt"
        var totalRow = worksheet.RowsUsed()
            .FirstOrDefault(row => row.CellsUsed().Any(cell => cell.GetValue<string>().Trim().Equals("I alt", StringComparison.OrdinalIgnoreCase)));

        if (totalRow == null)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("'I alt' row not found.");
            Console.ForegroundColor = ConsoleColor.Gray;
            return;
        }

        // Insert a new row above the "I alt" row
        int newRowNumber = totalRow.RowNumber();
        worksheet.Row(newRowNumber).InsertRowsAbove(1);
        var newRow = worksheet.Row(newRowNumber);

        // Find the column for "arrangementer"
        int arrangementerColumnIndex = GetColumnIndex(worksheet, "arrangementer");
        int fornavneColumnIndex = GetColumnIndex(worksheet, "Fornavne");

        if (arrangementerColumnIndex == -1)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("'arrangementer' column not found.");
            Console.ForegroundColor = ConsoleColor.Gray;
            return;
        }

        // Add the excluded amount to the new row
        newRow.Cell(arrangementerColumnIndex).Value = excludedAmount;
        newRow.Cell(fornavneColumnIndex).Value = $"{date} - {messages}";

        // Save the workbook
        workbook.Save();
        Console.WriteLine($"Excluded amount {excludedAmount} added to 'arrangementer' column above 'I alt'.");
    }
}

class Transaction
{
    public DateTime Date { get; set; }
    public required string Name { get; set; }
    public required string Phone { get; set; }
    public decimal Amount { get; set; }
    public string? Message { get; set; }
    public required string Type { get; set; }
    public required string TransactionID { get; set; }
    public bool IsExcluded { get; set; } 
}

internal class MonthlySummary
{
    public string Person { get; }
    public string Number { get; }
    public int Year { get; }
    public int Month { get; }
    public decimal Total { get; }

    public MonthlySummary(string person, string number, int year, int month, decimal total)
    {
        Person = person;
        Number = number;
        Year = year;
        Month = month;
        Total = total;
    }
}