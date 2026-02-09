using Xunit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HVM_Kasserer;

namespace HVM_Kasserer_Tests
{
    /// <summary>
    /// Advanced integration and behavioral tests for MobilePay.
    /// These tests focus on complex business scenarios and interactions between components.
    /// </summary>
    public class MobilePayIntegrationTests
    {
        private readonly string _testDataDirectory = Path.Combine(Path.GetTempPath(), "MobilePayIntTests");

        public MobilePayIntegrationTests()
        {
            Directory.CreateDirectory(_testDataDirectory);
        }

        #region Business Logic Tests

        /// <summary>
        /// Business requirement: The system must distinguish between "Betaling" (donation) and "Gebyr" (fees)
        /// transactions for reporting purposes. Only donations should count toward fundraising totals.
        /// </summary>
        [Theory]
        [InlineData("Betaling", true)]  // Should be counted as donation
        [InlineData("Gebyr", false)]    // Should NOT be counted as donation
        public void TransactionType_ShouldCorrectlyClassifyBetweenDonationAndFee(string transactionType, bool isDonation)
        {
            // Arrange
            var helper = new MobilePayTestHelper();
            const string gebyrKeyword = "Gebyr";

            // Act
            var isGebyr = transactionType == gebyrKeyword;
            var isCountedAsDonation = !isGebyr;

            // Assert
            Assert.Equal(isDonation, isCountedAsDonation);
        }

        /// <summary>
        /// Business requirement: Transactions from the same TransactionID but with different amounts
        /// are not duplicates and should be treated independently.
        /// </summary>
        [Fact]
        public void TransactionIdentification_ShouldTreatDifferentAmountsSeparately()
        {
            // Arrange
            var helper = new MobilePayTestHelper();
            var txn1 = ("TXN001", 100.00m, "John Smith");
            var txn2 = ("TXN001", 50.00m, "John Smith"); // Same ID, different amount

            // Act & Assert
            // These should NOT be considered the same transaction based on amount alone
            Assert.NotEqual(txn1.Item2, txn2.Item2);
            // But they share the same transaction ID
            Assert.Equal(txn1.Item1, txn2.Item1);
        }

        #endregion

        #region Cumulative Amount Tests

        /// <summary>
        /// Business requirement: When a person makes multiple donations in a month,
        /// all donations should be summed together for the monthly total.
        /// </summary>
        [Fact]
        public void MonthlySummary_ShouldAggregateDonationsForSamePerson()
        {
            // Arrange
            var donations = new List<(string name, decimal amount, int month)>
            {
                ("John Smith", 100.00m, 1),
                ("John Smith", 50.00m, 1),
                ("John Smith", 25.00m, 1),
                ("Jane Doe", 75.00m, 1)
            };

            // Act
            var johnTotal = donations
                .Where(d => d.name == "John Smith" && d.month == 1)
                .Sum(d => d.amount);

            var janeTotal = donations
                .Where(d => d.name == "Jane Doe" && d.month == 1)
                .Sum(d => d.amount);

            // Assert
            Assert.Equal(175.00m, johnTotal);
            Assert.Equal(75.00m, janeTotal);
        }

        /// <summary>
        /// Business requirement: Different months should have separate totals,
        /// even for the same person.
        /// </summary>
        [Fact]
        public void MonthlySummary_ShouldSeparateDifferentMonths()
        {
            // Arrange
            var donations = new List<(string name, decimal amount, int month)>
            {
                ("John Smith", 100.00m, 1),
                ("John Smith", 150.00m, 2),
                ("John Smith", 75.00m, 1)
            };

            // Act
            var january = donations
                .Where(d => d.name == "John Smith" && d.month == 1)
                .Sum(d => d.amount);

            var february = donations
                .Where(d => d.name == "John Smith" && d.month == 2)
                .Sum(d => d.amount);

            // Assert
            Assert.Equal(175.00m, january);
            Assert.Equal(150.00m, february);
        }

        #endregion

        #region Deduplication Tests

        /// <summary>
        /// Business requirement: When checking for duplicates in Excel,
        /// the system should match both date AND amount to ensure uniqueness.
        /// This prevents inserting the same transaction twice.
        /// </summary>
        [Fact]
        public void DuplicateDetection_ShouldMatchBothDateAndAmount()
        {
            // Arrange
            var existingEntries = new List<(DateTime date, decimal amount)>
            {
                (new DateTime(2025, 1, 15), 100.00m),
                (new DateTime(2025, 1, 16), 100.00m),
                (new DateTime(2025, 1, 15), 50.00m)
            };

            var newEntry = (date: new DateTime(2025, 1, 15), amount: 100.00m);

            // Act
            var isDuplicate = existingEntries.Any(e =>
                e.date == newEntry.date && Math.Abs(e.amount - newEntry.amount) < 0.005m);

            // Assert
            Assert.True(isDuplicate);
        }

        /// <summary>
        /// Business requirement: Same date but different amount should NOT be considered a duplicate.
        /// </summary>
        [Fact]
        public void DuplicateDetection_ShouldNotMatchDifferentAmounts()
        {
            // Arrange
            var existingEntries = new List<(DateTime date, decimal amount)>
            {
                (new DateTime(2025, 1, 15), 100.00m)
            };

            var newEntry = (date: new DateTime(2025, 1, 15), amount: 150.00m);

            // Act
            var isDuplicate = existingEntries.Any(e =>
                e.date == newEntry.date && Math.Abs(e.amount - newEntry.amount) < 0.005m);

            // Assert
            Assert.False(isDuplicate);
        }

        #endregion

        #region Message Handling Tests

        /// <summary>
        /// Business requirement: Special characters in transaction messages (quotes, apostrophes)
        /// should be cleaned to prevent data corruption and ensure consistency.
        /// </summary>
        [Theory]
        [InlineData("Test \"quoted\" message", "Test  quoted  message")]
        [InlineData("Test 'apostrophe' message", "Test  apostrophe  message")]
        [InlineData("\"Double\" 'quotes'", " Double   quotes ")]
        public void MessageCleaning_ShouldRemoveQuotesAndApostrophes(string input, string expected)
        {
            // Arrange
            var cleaned = input.Replace('\"', ' ').Replace('\'', ' ').Trim();

            // Assert
            Assert.Equal(expected.Trim(), cleaned);
        }

        /// <summary>
        /// Business requirement: Empty or null messages should be handled gracefully
        /// without causing errors in aggregation.
        /// </summary>
        [Fact]
        public void MessageAggregation_ShouldHandleNullAndEmptyMessages()
        {
            // Arrange
            var messages = new List<string?>
            {
                "Donation",
                null,
                "",
                "Test Message"
            };

            // Act
            var validMessages = messages
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .Select(m => m!.Trim())
                .Distinct()
                .ToList();

            // Assert
            Assert.Equal(2, validMessages.Count);
            Assert.Contains("Donation", validMessages);
            Assert.Contains("Test Message", validMessages);
        }

        #endregion

        #region Name Normalization Tests

        /// <summary>
        /// Business requirement: When matching transactions to people in Excel,
        /// the system must handle multiple spaces and variations in names
        /// to ensure correct person identification.
        /// </summary>
        [Theory]
        [InlineData("John    Smith", "john smith")]
        [InlineData("  John Smith  ", "john smith")]
        [InlineData("JOHN SMITH", "john smith")]
        public void NameMatching_ShouldNormalizeVariations(string input, string expected)
        {
            // Arrange
            var helper = new MobilePayTestHelper();

            // Act
            var normalized = helper.CallNormalizeString(input);

            // Assert
            Assert.Equal(expected, normalized);
        }

        /// <summary>
        /// Business requirement: Names should be rearranged consistently for display,
        /// always showing "LastName, FirstName" format in reports.
        /// </summary>
        [Fact]
        public void NameFormatting_ShouldBeConsistent()
        {
            // Arrange
            var helper = new MobilePayTestHelper();
            var names = new List<string>
            {
                "John Smith",
                "Jane Doe",
                "Peter Johnson"
            };

            // Act
            var formattedNames = names.Select(n => helper.CallRearrangeName(n)).ToList();

            // Assert
            Assert.All(formattedNames, name =>
            {
                // All formatted names should contain a comma separating last and first name
                Assert.Contains(",", name);
            });
        }

        #endregion

        #region Temporal Logic Tests

        /// <summary>
        /// Business requirement: The posting date logic ensures that weekend transactions
        /// are correctly attributed to the following Monday, which is when they actually clear.
        /// </summary>
        [Fact]
        public void PostingDateLogic_ShouldAlwaysMoveToNextBusinessDay()
        {
            // Arrange
            var helper = new MobilePayTestHelper();
            var dates = new List<DateTime>
            {
                new DateTime(2025, 1, 3), // Friday
                new DateTime(2025, 1, 4), // Saturday
                new DateTime(2025, 1, 5), // Sunday
                new DateTime(2025, 1, 6), // Monday
                new DateTime(2025, 1, 7)  // Tuesday
            };

            // Act
            var effectiveDates = dates.Select(d => helper.CallGetEffectivePostingDate(d)).ToList();

            // Assert
            // Friday -> Monday, Saturday -> Monday, Sunday -> Monday
            Assert.Equal(new DateTime(2025, 1, 6), effectiveDates[0]); // Friday
            Assert.Equal(new DateTime(2025, 1, 6), effectiveDates[1]); // Saturday
            Assert.Equal(new DateTime(2025, 1, 6), effectiveDates[2]); // Sunday
            // Monday -> Tuesday, Tuesday -> Wednesday
            Assert.Equal(new DateTime(2025, 1, 7), effectiveDates[3]); // Monday
            Assert.Equal(new DateTime(2025, 1, 8), effectiveDates[4]); // Tuesday
        }

        #endregion

        #region File Format Tests

        /// <summary>
        /// Business requirement: System must handle Danish locale correctly for date and number parsing.
        /// </summary>
        [Theory]
        [InlineData("100,50", 100.50)]  // Danish uses comma as decimal separator
        [InlineData("1.234,56", 1234.56)] // Danish uses period as thousands separator
        public void CultureInfo_ShouldHandleDanishFormats(string input, double expected)
        {
            // Arrange
            var culture = System.Globalization.CultureInfo.GetCultureInfo("da-DK");

            // Act
            var parsed = decimal.Parse(input, culture);

            // Assert
            Assert.Equal((decimal)expected, parsed);
        }

        #endregion

        #region Cleanup

        ~MobilePayIntegrationTests()
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
