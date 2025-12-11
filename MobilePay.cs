using ClosedXML.Excel;
using System.Globalization;
using System.Text.RegularExpressions;
using System.IO;

namespace HVM_Kasserer
{
    class MobilePay
    {
        private const string Gebyr = "Gebyr";
        private const string ColumnHeaderFornavne = "Fornavne";
        private const string ColumnHeaderMobil = "Mobil nummer";
        private const string ColumnHeaderArrangementer = "Arrangementer";
        private const string RowValueTotal = "I alt";
        private const string SheetNameMobilePay = "Mobilepay";

        static string basePath = @"C:\Dropbox\HVM - Kasserer";
        static string indsamlingerFolder = basePath + @"\Indsamlinger\2025 Indsamlinger";
        string mobilePayFilepath = indsamlingerFolder + @"\transactions-report.csv";
        string excelFilepath = indsamlingerFolder + @"\2025 Mobilepay.xlsx";
        string exclusionsFilePath = basePath + @"\Program-kode\HVM Kasserer\mobilePayExclusions.txt";
        string dailyReportsFolder = Path.Combine(indsamlingerFolder, "DailyReports");

        List<string> mobilePayExclusions;

        public MobilePay()
        {
            mobilePayExclusions = LoadExclusionsFromFile();
        }

        public void SummarizeMobilePayTransactions()
        {
            var transactions = GetTransactions();

            // Step 1: Identify TransactionIDs with exclusion keywords
            var excludedTransactionIDs = transactions
                .Where(t => t.Message != null &&
                            mobilePayExclusions.Any(keyword => NormalizeString(t.Message).Contains(NormalizeString(keyword))))
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
                .GroupBy(t => GetEffectivePostingDate(t.Date))
                .Select(group =>
                {
                    var regularTransactions = group.Where(t => !t.IsExcluded).ToList();
                    var excludedTransactions = group.Where(t => t.IsExcluded).ToList();

                    return new
                    {
                        Date = group.Key,
                        CombinedTotal = group.Sum(t => t.Amount),

                        RegularDonation = regularTransactions.Where(x => x.Type != Gebyr).Sum(x => x.Amount),
                        RegularTotal = regularTransactions.Sum(t => t.Amount),
                        RegularGebyr = regularTransactions.Where(x => x.Type == Gebyr).Sum(x => x.Amount) * -1,
                        RegularMessages = string.Join(", ", regularTransactions.Where(t => t.Message != null && t.Message.Length > 0).Select(t => t.Message).Distinct()),

                        ExcludedDonation = excludedTransactions.Where(x => x.Type != Gebyr).Sum(x => x.Amount),
                        ExcludedTotal = excludedTransactions.Sum(t => t.Amount),
                        ExcludedGebyr = excludedTransactions.Where(x => x.Type == Gebyr).Sum(x => x.Amount) * -1,
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

            Directory.CreateDirectory(dailyReportsFolder);

            var transactionsByDay = transactions.GroupBy(t => new { PostingDateTime = GetEffectivePostingDate(t.Date), t.Date.DayOfYear});
            foreach (var group in transactionsByDay)
            {
                bool multipleTransactions = transactionsByDay.Where(g => g.Key.PostingDateTime == group.Key.PostingDateTime).Count() > 1;
                WriteDailyTransactionsToExcel(group.Key.PostingDateTime, group.ToList(), multipleTransactions);
            }
            
            // Summarize by month for each person (excluding marked transactions)
            var monthlySummary = transactions
                .Where(t => !t.IsExcluded && t.Type != Gebyr)
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

                    UpdateExcelFile(SheetNameMobilePay, entry.Person, entry.Number, monthString, entry.Total);
                }

                Console.WriteLine($"");
            }

            // Replace excluded transactions grouping before UpdateExcelForExcluded
            var excludedTransactions = transactions.Where(t => t.IsExcluded && t.Type != Gebyr);
            var groupedByDay = excludedTransactions.GroupBy(t => GetEffectivePostingDate(t.Date));
            foreach (var date in groupedByDay)
            {
                var todaysTotal = date.Select(t => t.Amount).Sum();
                string dateString = date.Select(x => x.Date).First().ToString("yyyy-MM-dd");
                string? message = date.Where(x => x.Message != null).Select(x => x.Message).First();
                UpdateExcelForExcluded(SheetNameMobilePay, todaysTotal, dateString, message);
            }
        }

        private void UpdateExcelFile(string sheetName, string name, string last4PhoneDigits, string month, decimal value)
        {
            // Open the workbook
            using var workbook = new XLWorkbook(excelFilepath);
            var worksheet = workbook.Worksheet(sheetName);

            // Define column indices
            int nameColumnIndex = ExcelHelperMethods.GetColumnIndex(worksheet, ColumnHeaderFornavne);
            int phoneColumnIndex = ExcelHelperMethods.GetColumnIndex(worksheet, ColumnHeaderMobil);

            if (nameColumnIndex == -1 || phoneColumnIndex == -1)
            {
                Console.WriteLine($"Required columns ('{ColumnHeaderFornavne}' or '{ColumnHeaderMobil}') not found.");
                return;
            }

            // Find all rows matching the phone number
            var matchingRowsByPhone = worksheet.RowsUsed()
                .Where(row =>
                {
                    string rawPhoneNumber = row.Cell(phoneColumnIndex).GetValue<string>().Trim();
                    string cleanedPhoneNumber = HelperMethods.ExtractLast4Digits(rawPhoneNumber);
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
                    string existingPhoneNumber = HelperMethods.ExtractLast4Digits(personRow.Cell(phoneColumnIndex).GetValue<string>().Trim());

                    if (string.IsNullOrEmpty(existingPhoneNumber))
                    {
                        // Assign the phone number if it doesn't already exist
                        personRow.Cell(phoneColumnIndex).Value = last4PhoneDigits;
                    }
                    else if (existingPhoneNumber != last4PhoneDigits)
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
                // Find the last row that contains a value in the phone column and insert a new row after it.
                var rowsWithPhone = worksheet.RowsUsed()
                    .Where(r => !string.IsNullOrWhiteSpace(r.Cell(phoneColumnIndex).GetValue<string>()))
                    .ToList();

                int insertRowNumber;
                if (rowsWithPhone.Any())
                {
                    insertRowNumber = rowsWithPhone.Max(r => r.RowNumber()) + 1;
                }
                else
                {
                    // If no rows have a phone value, insert after the header (or at row 2 if the sheet is empty)
                    var headerRow = worksheet.FirstRowUsed();
                    insertRowNumber = (headerRow?.RowNumber() ?? 1) + 1;
                }

                // Insert a new row at the calculated position
                worksheet.Row(insertRowNumber).InsertRowsAbove(1);
                var newRow = worksheet.Row(insertRowNumber);

                newRow.Cell(nameColumnIndex).Value = name;
                newRow.Cell(phoneColumnIndex).Value = last4PhoneDigits;

                personRow = newRow;

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Added new person: {name} (phone: {last4PhoneDigits}) at row {insertRowNumber}.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            // Find the column for the month
            int monthColumnIndex = ExcelHelperMethods.GetColumnIndex(worksheet, month);

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

        private List<string> LoadExclusionsFromFile()
        {
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

        private string? GetMonthAsString(int monthNumber)
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

        private List<Transaction> GetTransactions()
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
                        Phone = HelperMethods.ExtractLast4Digits(values[16]),
                        Amount = decimal.Parse(values[6], cultureInfo),
                        Message = values[11].Replace('\"', ' ').Replace('\'', ' ').Trim(),
                        Type = values[5],
                        TransactionID = values[14],
                    });
                }
            }

            return transactionData;
        }

        private string NormalizeString(string input)
        {
            return Regex.Replace(input.ToLower(), @"\s+", " ").Trim();
        }

        private string RearrangeName(string fullName)
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

        // Method to insert excluded transactions above the "I alt" line
        private void UpdateExcelForExcluded(string sheetName, decimal excludedAmount, string date, string? messages)
        {
            using var workbook = new XLWorkbook(excelFilepath);
            var worksheet = workbook.Worksheet(sheetName);

            // Find the row that says "I alt"
            var totalRow = worksheet.RowsUsed()
                .FirstOrDefault(row => row.CellsUsed().Any(cell => cell.GetValue<string>().Trim().Equals(RowValueTotal, StringComparison.OrdinalIgnoreCase)));

            if (totalRow == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"'{RowValueTotal}' row not found.");
                Console.ForegroundColor = ConsoleColor.Gray;
                return;
            }

            // Find the column for "arrangementer"
            int arrangementerColumnIndex = ExcelHelperMethods.GetColumnIndex(worksheet, ColumnHeaderArrangementer);
            int fornavneColumnIndex = ExcelHelperMethods.GetColumnIndex(worksheet, ColumnHeaderFornavne);

            if (arrangementerColumnIndex == -1)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"'{ColumnHeaderArrangementer}' column not found.");
                Console.ForegroundColor = ConsoleColor.Gray;
                return;
            }

            if (fornavneColumnIndex == -1)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"'{ColumnHeaderFornavne}' column not found.");
                Console.ForegroundColor = ConsoleColor.Gray;
                return;
            }

            // Prepare the text we expect to find in the Fornavne column for this excluded entry
            string targetText = $"{date} - {messages ?? string.Empty}".Trim();

            // Check existing rows above the totalRow to avoid adding duplicates
            var rowsAbove = worksheet.RowsUsed().Where(r => r.RowNumber() < totalRow.RowNumber());
            foreach (var row in rowsAbove)
            {
                var nameCellText = row.Cell(fornavneColumnIndex).GetValue<string>()?.Trim() ?? string.Empty;
                if (!NormalizeString(nameCellText).Equals(NormalizeString(targetText), StringComparison.Ordinal))
                    continue;

                // Try to read the amount cell as a double; if it can't be parsed skip this row
                double cellAmount;
                try
                {
                    cellAmount = row.Cell(arrangementerColumnIndex).GetValue<double>();
                }
                catch
                {
                    continue;
                }

                if (Math.Abs(cellAmount - (double)excludedAmount) < 0.005)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Excluded transaction already present for {date} with amount {excludedAmount}. Skipping insertion.");
                    Console.ForegroundColor = ConsoleColor.Gray;
                    return;
                }
            }

            // Insert a new row above the "I alt" row
            int newRowNumber = totalRow.RowNumber();
            worksheet.Row(newRowNumber).InsertRowsAbove(1);
            var newRow = worksheet.Row(newRowNumber);

            // Add the excluded amount to the new row
            newRow.Cell(arrangementerColumnIndex).Value = excludedAmount;
            newRow.Cell(fornavneColumnIndex).Value = targetText;

            // Save the workbook
            workbook.Save();
            Console.WriteLine($"Excluded amount {excludedAmount} added to '{ColumnHeaderArrangementer}' column above '{RowValueTotal}'.");
        }

        private void WriteDailyTransactionsToExcel(DateTime date, List<Transaction> dayTransactions, bool multipleFiles)
        {
            string fileName = $"Mobilepay-{date:yyyy-MM-dd}";

            if (multipleFiles)
            {
                var sum = Math.Floor(dayTransactions.Sum(t => t.Amount));
                fileName = fileName + $"-{sum}";
            }

            if (dayTransactions.Any(dayTransactions => dayTransactions.IsExcluded))
                fileName = fileName + "-hasExcluded";

            var fullFileName = fileName + ".xlsx";

            string filePath = Path.Combine(dailyReportsFolder, fullFileName);

            using var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Transactions");

            // Headers
            var headers = new[] { "Date", "Time", "Name", "Phone", "Type", "Amount", "Message", "TransactionID", "Gift" };
            for (int i = 0; i < headers.Length; i++)
                ws.Cell(1, i + 1).Value = headers[i];

            int row = 2;
            foreach (var t in dayTransactions)
            {
                ws.Cell(row, 1).Value = t.Date.ToString("yyyy-MM-dd", CultureInfo.GetCultureInfo("da-DK"));
                ws.Cell(row, 2).Value = t.Date.ToString("HH:mm:ss");
                ws.Cell(row, 3).Value = t.Name;
                ws.Cell(row, 4).Value = t.Phone;
                ws.Cell(row, 5).Value = t.Type;
                ws.Cell(row, 6).Value = t.Amount;
                ws.Cell(row, 7).Value = t.Message;
                ws.Cell(row, 8).Value = t.TransactionID;
                ws.Cell(row, 9).Value = t.IsExcluded ? "No" : "Yes";
                row++;
            }

            // Sum row
            ws.Cell(row, 5).Value = "Total";
            ws.Cell(row, 6).FormulaA1 = $"=SUM(F2:F{row - 1})";

            // Formatting
            ws.Column(6).Style.NumberFormat.Format = "#,##0.00";
            ws.Columns().AdjustToContents();

            workbook.SaveAs(filePath);
            Console.WriteLine($"Saved daily file: {filePath}");
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

        // Add this helper inside the MobilePay class (near other private helpers)
        private DateTime GetEffectivePostingDate(DateTime dateTime)
        {
            // Transactions that hit on Friday, Saturday or Sunday should be treated as posted on the following Monday
            var date = dateTime.Date;
            return date.DayOfWeek switch
            {
                DayOfWeek.Friday => date.AddDays(3),
                DayOfWeek.Saturday => date.AddDays(2),
                _ => date.AddDays(1)
            };
        }
    }
}
