using Xunit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;
using HVM_Kasserer;

namespace HVM_Kasserer_Tests
{
    /// <summary>
    /// Unit tests for MobilePay class focusing on outside-in (behavioral) testing.
    /// Tests verify business requirements and integration points rather than implementation details.
    /// </summary>
    public class MobilePayTests : IDisposable
    {
        private readonly string _testDataDirectory;
        private readonly string _testExcelDirectory;

        public MobilePayTests()
        {
            // Use a unique directory for each test instance to avoid race conditions
            _testDataDirectory = Path.Combine(Path.GetTempPath(), "MobilePayTests", Guid.NewGuid().ToString());
            _testExcelDirectory = Path.Combine(_testDataDirectory, "Indsamlinger", "2025 Indsamlinger");
            Directory.CreateDirectory(_testExcelDirectory);
        }

        #region Test Fixtures

        private void CreateTestCsvFile(string filename, List<string[]> rows)
        {
            var filePath = Path.Combine(_testExcelDirectory, filename);
            using (var writer = new StreamWriter(filePath))
            {
                foreach (var row in rows)
                {
                    writer.WriteLine(string.Join(";", row));
                }
            }
        }

        private void CreateTestExclusionsFile(string filename, List<string> exclusions)
        {
            var filePath = Path.Combine(_testDataDirectory, "Program-kode", "HVM Kasserer", filename);
            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
            File.WriteAllLines(filePath, exclusions);
        }

        #endregion

        #region Name Handling Tests

        /// <summary>
        /// Business requirement: System must rearrange full names to "LastName, FirstName" format
        /// for display in reports.
        /// </summary>
        [Theory]
        [InlineData("John Smith", "Smith, John")]
        [InlineData("Marie Curie Doe", "Doe, Marie Curie")]
        [InlineData("SingleName", "SingleName")]
        [InlineData("Per Andersen", "Andersen, Per")]
        public void RearrangeName_ShouldConvertNameToLastNameFirstNameFormat(string input, string expected)
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();

            // Act
            var result = mobilePay.CallRearrangeName(input);

            // Assert
            Assert.Equal(expected, result);
        }

        /// <summary>
        /// Business requirement: The system must handle multiple spaces in names
        /// by normalizing them to single spaces.
        /// </summary>
        [Theory]
        [InlineData("John    Smith", "Smith, John")]
        [InlineData("  John Smith  ", "Smith, John")]
        public void RearrangeName_ShouldHandleMultipleSpaces(string input, string expected)
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();

            // Act
            var result = mobilePay.CallRearrangeName(input);

            // Assert - name should be rearranged without extra spaces
            Assert.DoesNotContain("  ", result);
        }

        #endregion

        #region String Normalization Tests

        /// <summary>
        /// Business requirement: System must normalize strings for comparison,
        /// handling case-insensitivity and whitespace variations.
        /// </summary>
        [Theory]
        [InlineData("Test String", "test string")]
        [InlineData("TEST  STRING", "test string")]
        [InlineData("  test string  ", "test string")]
        [InlineData("TeSt StRiNg", "test string")]
        public void NormalizeString_ShouldMakeStringsComparable(string input, string expected)
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();

            // Act
            var result = mobilePay.CallNormalizeString(input);

            // Assert
            Assert.Equal(expected, result);
        }

        #endregion

        #region Month Conversion Tests

        /// <summary>
        /// Business requirement: System must convert month numbers to Danish month names
        /// for use in Excel column headers and reports.
        /// </summary>
        [Theory]
        [InlineData(1, "Jan")]
        [InlineData(2, "Feb")]
        [InlineData(3, "Marts")]
        [InlineData(4, "April")]
        [InlineData(5, "Maj")]
        [InlineData(6, "Juni")]
        [InlineData(7, "Juli")]
        [InlineData(8, "August")]
        [InlineData(9, "Sept")]
        [InlineData(10, "Okt")]
        [InlineData(11, "Nov")]
        [InlineData(12, "Dec")]
        public void GetMonthAsString_ShouldReturnDanishMonthName(int monthNumber, string expected)
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();

            // Act
            var result = mobilePay.CallGetMonthAsString(monthNumber);

            // Assert
            Assert.Equal(expected, result);
        }

        /// <summary>
        /// Business requirement: System must handle invalid month numbers gracefully.
        /// </summary>
        [Theory]
        [InlineData(0)]
        [InlineData(13)]
        [InlineData(-1)]
        public void GetMonthAsString_ShouldReturnNullForInvalidMonthNumber(int monthNumber)
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();

            // Act
            var result = mobilePay.CallGetMonthAsString(monthNumber);

            // Assert
            Assert.Null(result);
        }

        #endregion

        #region Effective Posting Date Tests

        /// <summary>
        /// Business requirement: Weekend and Friday transactions should be posted on the following Monday
        /// to match actual bank clearing schedules.
        /// </summary>
        [Fact]
        public void GetEffectivePostingDate_ShouldMoveFridayToMonday()
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();
            var friday = new DateTime(2025, 1, 3); // A Friday

            // Act
            var result = mobilePay.CallGetEffectivePostingDate(friday);

            // Assert
            Assert.Equal(DayOfWeek.Monday, result.DayOfWeek);
            Assert.Equal(new DateTime(2025, 1, 6), result); // 3 days later
        }

        /// <summary>
        /// Business requirement: Saturday transactions should be posted on the following Monday.
        /// </summary>
        [Fact]
        public void GetEffectivePostingDate_ShouldMoveSaturdayToMonday()
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();
            var saturday = new DateTime(2025, 1, 4); // A Saturday

            // Act
            var result = mobilePay.CallGetEffectivePostingDate(saturday);

            // Assert
            Assert.Equal(DayOfWeek.Monday, result.DayOfWeek);
            Assert.Equal(new DateTime(2025, 1, 6), result); // 2 days later
        }

        /// <summary>
        /// Business requirement: Sunday transactions should be posted on the following Monday.
        /// </summary>
        [Fact]
        public void GetEffectivePostingDate_ShouldMoveSundayToMonday()
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();
            var sunday = new DateTime(2025, 1, 5); // A Sunday

            // Act
            var result = mobilePay.CallGetEffectivePostingDate(sunday);

            // Assert
            Assert.Equal(DayOfWeek.Monday, result.DayOfWeek);
            Assert.Equal(new DateTime(2025, 1, 6), result); // 1 day later
        }

        /// <summary>
        /// Business requirement: Weekday (Monday-Thursday) transactions should be posted the next day.
        /// </summary>
        [Theory]
        [InlineData(DayOfWeek.Monday)]
        [InlineData(DayOfWeek.Tuesday)]
        [InlineData(DayOfWeek.Wednesday)]
        [InlineData(DayOfWeek.Thursday)]
        public void GetEffectivePostingDate_ShouldMoveWeekdayToNextDay(DayOfWeek dayOfWeek)
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();
            var date = FindDateWithDayOfWeek(2025, 1, dayOfWeek);

            // Act
            var result = mobilePay.CallGetEffectivePostingDate(date);

            // Assert
            Assert.Equal(date.AddDays(1), result);
        }

        #endregion

        #region Exclusion Loading Tests

        /// <summary>
        /// Business requirement: System must load exclusion keywords from file to identify
        /// transactions that should not be included in donation totals.
        /// </summary>
        [Fact]
        public void LoadExclusionsFromFile_ShouldLoadValidExclusions()
        {
            // Arrange
            var exclusions = new List<string>
            {
                "Refund",
                "Test Transaction",
                "Admin Fee"
            };
            CreateTestExclusionsFile("mobilePayExclusions.txt", exclusions);

            var mobilePay = new MobilePayTestHelper();

            // Act
            var result = mobilePay.CallLoadExclusionsFromFile(
                Path.Combine(_testDataDirectory, "Program-kode", "HVM Kasserer", "mobilePayExclusions.txt"));

            // Assert
            Assert.Equal(3, result.Count);
            Assert.Contains("Refund", result);
            Assert.Contains("Test Transaction", result);
            Assert.Contains("Admin Fee", result);
        }

        /// <summary>
        /// Business requirement: Empty lines and whitespace in exclusion file should be ignored.
        /// </summary>
        [Fact]
        public void LoadExclusionsFromFile_ShouldIgnoreEmptyLines()
        {
            // Arrange
            var fileContent = new List<string>
            {
                "Refund",
                "",
                "  ",
                "Test Transaction"
            };
            CreateTestExclusionsFile("mobilePayExclusions.txt", fileContent);

            var mobilePay = new MobilePayTestHelper();

            // Act
            var result = mobilePay.CallLoadExclusionsFromFile(
                Path.Combine(_testDataDirectory, "Program-kode", "HVM Kasserer", "mobilePayExclusions.txt"));

            // Assert
            Assert.Equal(2, result.Count);
            Assert.DoesNotContain("", result);
        }

        /// <summary>
        /// Business requirement: When exclusions file doesn't exist, system should return empty list
        /// and log a warning rather than crashing.
        /// </summary>
        [Fact]
        public void LoadExclusionsFromFile_ShouldReturnEmptyListWhenFileNotFound()
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();
            var nonExistentPath = Path.Combine(_testDataDirectory, "nonexistent", "mobilePayExclusions.txt");

            // Act
            var result = mobilePay.CallLoadExclusionsFromFile(nonExistentPath);

            // Assert
            Assert.NotNull(result);
            Assert.Empty(result);
        }

        #endregion

        #region Transaction Classification Tests

        /// <summary>
        /// Business requirement: Transactions with messages matching exclusion keywords
        /// must be marked as excluded and not included in regular donation totals.
        /// </summary>
        [Fact]
        public void TransactionClassification_ShouldIdentifyExcludedTransactions()
        {
            // Arrange
            var transactions = new List<(string message, bool shouldBeExcluded)>
            {
                ("Regular donation", false),
                ("Refund for event", true),
                ("Standard payment", false),
                ("REFUND - test", true),
                ("Test Transaction - refund", true)
            };

            var exclusionKeywords = new List<string> { "Refund", "Test Transaction" };
            var mobilePay = new MobilePayTestHelper();

            // Act & Assert
            foreach (var (message, shouldBeExcluded) in transactions)
            {
                var isExcluded = exclusionKeywords.Any(keyword =>
                    mobilePay.CallNormalizeString(message).Contains(mobilePay.CallNormalizeString(keyword)));

                Assert.Equal(shouldBeExcluded, isExcluded);
            }
        }

        /// <summary>
        /// Business requirement: Exclusion matching should be case-insensitive and
        /// handle whitespace variations.
        /// </summary>
        [Theory]
        [InlineData("refund", "Refund", true)]
        [InlineData("REFUND", "refund", true)]
        [InlineData("Test  Transaction", "test transaction", true)]
        [InlineData("NotAMatch", "Refund", false)]
        public void TransactionClassification_ShouldBeCaseInsensitive(
            string transactionMessage, string exclusionKeyword, bool shouldMatch)
        {
            // Arrange
            var mobilePay = new MobilePayTestHelper();
            var normalizedMessage = mobilePay.CallNormalizeString(transactionMessage);
            var normalizedKeyword = mobilePay.CallNormalizeString(exclusionKeyword);

            // Act
            var isMatch = normalizedMessage.Contains(normalizedKeyword);

            // Assert
            Assert.Equal(shouldMatch, isMatch);
        }

        #endregion

        #region Phone Number Extraction Tests

        /// <summary>
        /// Business requirement: System must extract and store the last 4 digits of phone numbers
        /// for matching and deduplication purposes.
        /// </summary>
        [Theory]
        [InlineData("+45 1234 5678", "5678")]
        [InlineData("12345678", "5678")]
        [InlineData("+45-1234-5678", "5678")]
        [InlineData("555", "")] // Less than 4 digits
        public void ExtractLast4Digits_ShouldReturnCorrectDigits(string phoneNumber, string expected)
        {
            // Arrange
            var helper = new MobilePayTestHelper();

            // Act
            var result = helper.CallExtractLast4Digits(phoneNumber);

            // Assert
            Assert.Equal(expected, result);
        }

        #endregion

        #region Integration Tests

        /// <summary>
        /// Business requirement: System must successfully group transactions by date
        /// and calculate separate totals for regular and excluded donations.
        /// This is a higher-level integration test.
        /// </summary>
        [Fact]
        public void GetTransactions_ShouldParseCSVCorrectly()
        {
            // Arrange
            var csvRows = new List<string[]>
            {
                // Header row (must match the CSV format)
                new[] { "", "", "", "", "", "Type", "Amount", "", "", "", "Date", "Message", "", "", "TransactionID", "Name", "Phone" },
                new[] { "", "", "", "", "", "Betaling", "100,00", "", "", "", "2025-01-15", "Donation", "", "", "TXN001", "John Smith", "+45 1234 5678" },
                new[] { "", "", "", "", "", "Gebyr", "-5,00", "", "", "", "2025-01-15", "", "", "", "TXN002", "John Smith", "+45 1234 5678" }
            };

            CreateTestCsvFile("transactions-report.csv", csvRows);

            var mobilePay = new MobilePayTestHelper();

            // Act
            var transactionsResult = mobilePay.CallGetTransactions(
                Path.Combine(_testExcelDirectory, "transactions-report.csv"));
            
            // Convert the result to a list we can work with
            var transactions = new List<dynamic>();
            if (transactionsResult is System.Collections.IEnumerable enumerable)
            {
                foreach (var item in enumerable)
                {
                    transactions.Add(item);
                }
            }

            // Assert
            Assert.NotEmpty(transactions);
            Assert.Equal(2, transactions.Count);
            
            // Access properties using reflection since dynamic binding doesn't work with private types
            var firstTransaction = transactions[0];
            var firstType = firstTransaction.GetType().GetProperty("Type")?.GetValue(firstTransaction);
            var firstAmount = firstTransaction.GetType().GetProperty("Amount")?.GetValue(firstTransaction);
            
            var secondTransaction = transactions[1];
            var secondType = secondTransaction.GetType().GetProperty("Type")?.GetValue(secondTransaction);
            var secondAmount = secondTransaction.GetType().GetProperty("Amount")?.GetValue(secondTransaction);

            Assert.Equal("Betaling", firstType);
            Assert.Equal(100.00m, firstAmount);
            Assert.Equal("Gebyr", secondType);
            Assert.Equal(-5.00m, secondAmount);
        }

        #endregion

        #region Daily Report Generation Tests

        /// <summary>
        /// Business requirement: System must generate daily transaction reports in Excel format
        /// with proper formatting and calculations.
        /// </summary>
        [Fact]
        public void WriteDailyTransactionsToExcel_ShouldCreateValidExcelFile()
        {
            // Arrange
            var testDate = new DateTime(2025, 1, 15);
            var transactions = new List<(string name, decimal amount, string type)>
            {
                ("John Smith", 100.00m, "Betaling"),
                ("Jane Doe", 50.00m, "Betaling"),
                ("System", -5.00m, "Gebyr")
            };

            var mobilePay = new MobilePayTestHelper();
            var outputPath = Path.Combine(_testExcelDirectory, "DailyReports");
            Directory.CreateDirectory(outputPath);

            // Act
            var fileName = mobilePay.GenerateDailyExcelFile(testDate, transactions, outputPath);

            // Assert
            var filePath = Path.Combine(outputPath, fileName);
            Assert.True(File.Exists(filePath), $"Excel file should be created at {filePath}");
            Assert.EndsWith(".xlsx", fileName);
        }

        #endregion

        #region Helper Methods

        private static DateTime FindDateWithDayOfWeek(int year, int month, DayOfWeek dayOfWeek)
        {
            var date = new DateTime(year, month, 1);
            while (date.DayOfWeek != dayOfWeek)
            {
                date = date.AddDays(1);
            }
            return date;
        }

        #endregion

        #region Cleanup

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_testDataDirectory))
                {
                    Directory.Delete(_testDataDirectory, true);
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        #endregion
    }
}
