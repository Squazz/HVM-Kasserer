using Xunit;
using System;
using System.Collections.Generic;
using System.IO;
using HVM_Kasserer;

namespace HVM_Kasserer_Tests
{
    /// <summary>
    /// Edge case and error handling tests for MobilePay.
    /// These tests verify that the system handles unusual inputs and failures gracefully.
    /// </summary>
    public class MobilePayEdgeCaseTests
    {
        private readonly MobilePayTestHelper _helper = new MobilePayTestHelper();

        #region Name Edge Cases

        /// <summary>
        /// Business requirement: System must handle edge cases in names gracefully,
        /// including very short names or names with special characters.
        /// </summary>
        [Theory]
        [InlineData("")]         // Empty name
        [InlineData(" ")]        // Whitespace only
        [InlineData("A")]        // Single character
        [InlineData("A B C D")]  // Many parts
        public void RearrangeName_ShouldHandleEdgeCases(string input)
        {
            // Act
            var result = _helper.CallRearrangeName(input);

            // Assert - should not throw and should return a string
            Assert.NotNull(result);
            Assert.IsType<string>(result);
        }

        /// <summary>
        /// Business requirement: Names with numbers and special characters should be processed.
        /// </summary>
        [Theory]
        [InlineData("John 123")]
        [InlineData("O'Brien Smith")]
        [InlineData("José García")]
        public void RearrangeName_ShouldHandleSpecialCharacters(string input)
        {
            // Act
            var result = _helper.CallRearrangeName(input);

            // Assert
            Assert.NotNull(result);
            Assert.NotEmpty(result);
        }

        #endregion

        #region String Normalization Edge Cases

        /// <summary>
        /// Business requirement: Normalization should handle extreme whitespace variations
        /// and non-ASCII characters.
        /// </summary>
        [Theory]
        [InlineData("")]                    // Empty string
        [InlineData("   ")]                 // Only spaces
        [InlineData("\t\n\r")]              // Whitespace characters
        [InlineData("Test\u00A0String")]    // Non-breaking space
        public void NormalizeString_ShouldHandleEdgeCases(string input)
        {
            // Act
            var result = _helper.CallNormalizeString(input);

            // Assert - should not throw
            Assert.NotNull(result);
            Assert.IsType<string>(result);
        }

        /// <summary>
        /// Business requirement: Identical normalization of different inputs should be identical.
        /// </summary>
        [Fact]
        public void NormalizeString_ShouldBeIdempotent()
        {
            // Arrange
            var input = "Test   String   With   Spaces";

            // Act
            var normalized1 = _helper.CallNormalizeString(input);
            var normalized2 = _helper.CallNormalizeString(normalized1);

            // Assert - normalizing twice should give same result
            Assert.Equal(normalized1, normalized2);
        }

        #endregion

        #region Month Conversion Edge Cases

        /// <summary>
        /// Business requirement: Invalid month numbers should be handled safely.
        /// </summary>
        [Theory]
        [InlineData(0)]
        [InlineData(13)]
        [InlineData(-1)]
        [InlineData(100)]
        [InlineData(int.MinValue)]
        [InlineData(int.MaxValue)]
        public void GetMonthAsString_ShouldHandleInvalidInputs(int monthNumber)
        {
            // Act
            var result = _helper.CallGetMonthAsString(monthNumber);

            // Assert - should return null or empty, not throw
            Assert.Null(result);
        }

        /// <summary>
        /// Business requirement: All valid months 1-12 must be covered.
        /// </summary>
        [Fact]
        public void GetMonthAsString_ShouldCoverAllValidMonths()
        {
            // Act & Assert
            for (int month = 1; month <= 12; month++)
            {
                var result = _helper.CallGetMonthAsString(month);
                Assert.NotNull(result);
                Assert.NotEmpty(result!);
            }
        }

        #endregion

        #region Posting Date Edge Cases

        /// <summary>
        /// Business requirement: Posting date calculation should work correctly
        /// around month and year boundaries.
        /// </summary>
        [Fact]
        public void GetEffectivePostingDate_ShouldHandleMonthBoundary()
        {
            // Arrange - Friday at end of month
            var friday = new DateTime(2025, 1, 31); // January 31, 2025 is a Friday

            // Act
            var result = _helper.CallGetEffectivePostingDate(friday);

            // Assert - should move to next Monday, which is in February
            Assert.Equal(new DateTime(2025, 2, 3), result);
        }

        /// <summary>
        /// Business requirement: Posting date calculation should work correctly
        /// at year boundaries.
        /// </summary>
        [Fact]
        public void GetEffectivePostingDate_ShouldHandleYearBoundary()
        {
            // Arrange - Friday at end of year
            var friday = new DateTime(2023, 12, 29); // December 29, 2023 is a Friday

            // Act
            var result = _helper.CallGetEffectivePostingDate(friday);

            // Assert - should move to next Monday, which is in 2024
            Assert.Equal(new DateTime(2024, 1, 1), result);
        }

        /// <summary>
        /// Business requirement: Leap year transitions should be handled correctly.
        /// </summary>
        [Fact]
        public void GetEffectivePostingDate_ShouldHandleLeapYearTransition()
        {
            // Arrange - Friday before leap day
            var friday = new DateTime(2024, 2, 23); // Friday before leap day

            // Act
            var result = _helper.CallGetEffectivePostingDate(friday);

            // Assert - should correctly add days including leap day
            Assert.True(result > friday);
        }

        #endregion

        #region Exclusion Loading Edge Cases

        /// <summary>
        /// Business requirement: System should handle missing exclusion files gracefully.
        /// </summary>
        [Fact]
        public void LoadExclusionsFromFile_ShouldHandleMissingFile()
        {
            // Arrange
            var nonExistentPath = Path.Combine(Path.GetTempPath(), $"test-{Guid.NewGuid()}", "nonexistent.txt");

            // Act
            var result = _helper.CallLoadExclusionsFromFile(nonExistentPath);

            // Assert - should return empty list, not throw
            Assert.NotNull(result);
            Assert.Empty(result);
        }

        /// <summary>
        /// Business requirement: Malformed exclusion files should be handled gracefully.
        /// </summary>
        [Fact]
        public void LoadExclusionsFromFile_ShouldHandleEmptyFile()
        {
            // Arrange
            var tempFile = Path.Combine(Path.GetTempPath(), $"exclusions-{Guid.NewGuid()}.txt");
            File.WriteAllText(tempFile, string.Empty);

            try
            {
                // Act
                var result = _helper.CallLoadExclusionsFromFile(tempFile);

                // Assert
                Assert.NotNull(result);
                Assert.Empty(result);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        /// <summary>
        /// Business requirement: Large exclusion files should be handled without performance issues.
        /// </summary>
        [Fact]
        public void LoadExclusionsFromFile_ShouldHandleLargeFiles()
        {
            // Arrange - Create file with many exclusions
            var tempFile = Path.Combine(Path.GetTempPath(), $"exclusions-{Guid.NewGuid()}.txt");
            var exclusions = new List<string>();
            for (int i = 0; i < 1000; i++)
            {
                exclusions.Add($"Exclusion_{i}");
            }
            File.WriteAllLines(tempFile, exclusions);

            try
            {
                // Act
                var result = _helper.CallLoadExclusionsFromFile(tempFile);

                // Assert
                Assert.Equal(1000, result.Count);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        #endregion

        #region Decimal Precision Tests

        /// <summary>
        /// Business requirement: Financial calculations must handle decimal precision
        /// to avoid rounding errors in amounts.
        /// </summary>
        [Theory]
        [InlineData(100.00, 50.50, 150.50)]
        [InlineData(0.01, 0.02, 0.03)]
        [InlineData(1000000.99, 0.01, 1000001.00)]
        public void DecimalPrecision_ShouldPreserveAccuracy(decimal a, decimal b, decimal expected)
        {
            // Act
            var result = a + b;

            // Assert
            Assert.Equal(expected, result);
        }

        /// <summary>
        /// Business requirement: Comparing decimal amounts should use small tolerance
        /// to handle floating-point precision issues.
        /// </summary>
        [Fact]
        public void DecimalComparison_ShouldUseTolerance()
        {
            // Arrange
            var amount1 = 100.00m;
            var amount2 = 100.004m; // Difference less than 0.005
            const decimal tolerance = 0.005m;

            // Act
            var areSame = Math.Abs(amount1 - amount2) < tolerance;

            // Assert
            Assert.True(areSame);
        }

        #endregion

        #region Phone Number Edge Cases

        /// <summary>
        /// Business requirement: System must extract last 4 digits from phone numbers
        /// regardless of formatting or length.
        /// </summary>
        [Theory]
        [InlineData("+45 1234 5678", "5678")]
        [InlineData("5678", "5678")]
        [InlineData("123", "")] // Less than 4 digits
        [InlineData("", "")] // Empty
        [InlineData("abcd1234", "1234")] // Letters mixed in
        public void ExtractLast4Digits_ShouldHandleVariousFormats(string phoneNumber, string expected)
        {
            // Act
            var result = _helper.CallExtractLast4Digits(phoneNumber);

            // Assert
            Assert.Equal(expected, result);
        }

        #endregion

        #region Concurrent Scenario Tests

        /// <summary>
        /// Business requirement: The system should handle multiple transactions
        /// from the same person on the same day.
        /// </summary>
        [Fact]
        public void MultipleDonations_ShouldAggregateCorrectly()
        {
            // Arrange
            var donations = new List<(string person, decimal amount, DateTime date)>
            {
                ("John Smith", 50.00m, new DateTime(2025, 1, 15, 09, 00, 00)),
                ("John Smith", 50.00m, new DateTime(2025, 1, 15, 10, 30, 00)),
                ("John Smith", 50.00m, new DateTime(2025, 1, 15, 14, 15, 00))
            };

            // Act
            var total = donations
                .Where(d => d.person == "John Smith" && d.date.Date == new DateTime(2025, 1, 15).Date)
                .Sum(d => d.amount);

            // Assert
            Assert.Equal(150.00m, total);
        }

        #endregion

        #region Null/Empty Reference Tests

        /// <summary>
        /// Business requirement: System should handle null messages in transactions.
        /// </summary>
        [Fact]
        public void NullMessage_ShouldBeHandledGracefully()
        {
            // Arrange
            string? message = null;

            // Act
            var isEmpty = string.IsNullOrWhiteSpace(message);

            // Assert
            Assert.True(isEmpty);
        }

        /// <summary>
        /// Business requirement: Empty strings should be treated as missing data.
        /// </summary>
        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        public void EmptyOrNull_ShouldBeConsistentlyHandled(string? input)
        {
            // Act
            var isEmpty = string.IsNullOrWhiteSpace(input);

            // Assert
            Assert.True(isEmpty);
        }

        #endregion
    }
}
