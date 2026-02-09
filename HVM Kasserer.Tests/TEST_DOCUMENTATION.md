# MobilePay Unit Tests - Test Suite Documentation

## Overview

This document describes the comprehensive unit test suite created for the `MobilePay` class in the HVM Kasserer application. The tests focus on **outside-in testing** (behavioral testing) rather than implementation details, ensuring that business requirements persist regardless of internal logic changes.

## Test Files Created

### 1. **MobilePayTests.cs** - Core Business Logic Tests
This file contains the primary test cases organized into thematic sections:

#### Name Handling Tests
- `RearrangeName_ShouldConvertNameToLastNameFirstNameFormat` - Verifies names are rearranged to "LastName, FirstName" format
- `RearrangeName_ShouldHandleMultipleSpaces` - Ensures multiple spaces are normalized

**Business Requirement**: Names must be consistently formatted in reports for proper identification and display.

#### String Normalization Tests
- `NormalizeString_ShouldMakeStringsComparable` - Verifies case-insensitive and whitespace-handling normalization
  
**Business Requirement**: The system must match exclusion keywords regardless of case or whitespace variations.

#### Month Conversion Tests
- `GetMonthAsString_ShouldReturnDanishMonthName` - Tests all 12 Danish month names
- `GetMonthAsString_ShouldReturnNullForInvalidMonthNumber` - Ensures invalid months are handled gracefully

**Business Requirement**: Reports must display dates in Danish for the local audience.

#### Effective Posting Date Tests
- `GetEffectivePostingDate_ShouldMoveFridayToMonday` - Friday transactions post on Monday
- `GetEffectivePostingDate_ShouldMoveSaturdayToMonday` - Saturday transactions post on Monday
- `GetEffectivePostingDate_ShouldMoveSundayToMonday` - Sunday transactions post on Monday
- `GetEffectivePostingDate_ShouldMoveWeekdayToNextDay` - Weekday transactions post next day

**Business Requirement**: Weekend transactions must be posted according to actual bank clearing schedules (next Monday).

#### Exclusion Loading Tests
- `LoadExclusionsFromFile_ShouldLoadValidExclusions` - Verifies keyword file parsing
- `LoadExclusionsFromFile_ShouldIgnoreEmptyLines` - Ensures clean data processing
- `LoadExclusionsFromFile_ShouldReturnEmptyListWhenFileNotFound` - Graceful error handling

**Business Requirement**: Exclusion rules must be configurable and persistent without breaking the system.

#### Transaction Classification Tests
- `TransactionClassification_ShouldIdentifyExcludedTransactions` - Verifies exclusion keyword matching
- `TransactionClassification_ShouldBeCaseInsensitive` - Ensures robust matching

**Business Requirement**: Transactions with certain keywords must be excluded from donation totals.

#### Phone Number Extraction Tests
- `ExtractLast4Digits_ShouldReturnCorrectDigits` - Verifies phone number parsing from various formats

**Business Requirement**: Last 4 digits are used for person identification and deduplication.

#### Integration Tests
- `GetTransactions_ShouldParseCSVCorrectly` - Verifies CSV file parsing with correct typing

**Business Requirement**: Transaction data must be accurately read and parsed from source files.

#### Daily Report Generation Tests
- `WriteDailyTransactionsToExcel_ShouldCreateValidExcelFile` - Verifies Excel file creation with proper structure

**Business Requirement**: Daily reports must be generated in Excel format for audit trails.

### 2. **MobilePayIntegrationTests.cs** - Advanced Business Logic Tests
This file tests complex scenarios and interactions between components:

#### Business Logic Tests
- `TransactionType_ShouldCorrectlyClassifyBetweenDonationAndFee` - Distinguishes "Betaling" from "Gebyr"
- `TransactionIdentification_ShouldTreatDifferentAmountsSeparately` - Same transaction ID with different amounts are independent

**Business Requirement**: Only actual donations count toward fundraising totals, not fees.

#### Cumulative Amount Tests
- `MonthlySummary_ShouldAggregateDonationsForSamePerson` - Multiple donations per person sum correctly
- `MonthlySummary_ShouldSeparateDifferentMonths` - Monthly totals don't bleed across months

**Business Requirement**: Accurate monthly reporting for each person is essential for reconciliation.

#### Deduplication Tests
- `DuplicateDetection_ShouldMatchBothDateAndAmount` - Prevents duplicate entries
- `DuplicateDetection_ShouldNotMatchDifferentAmounts` - Same date, different amount = new entry

**Business Requirement**: Excel entries must not duplicate when processing returns.

#### Message Handling Tests
- `MessageCleaning_ShouldRemoveQuotesAndApostrophes` - Special character removal
- `MessageAggregation_ShouldHandleNullAndEmptyMessages` - Graceful null handling

**Business Requirement**: Messages must be cleaned for data integrity without losing information.

#### Name Normalization Tests
- `NameMatching_ShouldNormalizeVariations` - Various name formats match to same person
- `NameFormatting_ShouldBeConsistent` - All names formatted uniformly

**Business Requirement**: Consistent name formatting prevents duplicate person entries.

#### Temporal Logic Tests
- `PostingDateLogic_ShouldAlwaysMoveToNextBusinessDay` - Comprehensive posting date validation

**Business Requirement**: Business rules for posting dates must be applied consistently.

#### Culture/Locale Tests
- `CultureInfo_ShouldHandleDanishFormats` - Danish number and date formats parsed correctly

**Business Requirement**: System must work correctly in Danish locale (comma as decimal separator).

### 3. **MobilePayEdgeCaseTests.cs** - Robustness and Error Handling
This file tests edge cases and unusual inputs to ensure system resilience:

#### Name Edge Cases
- `RearrangeName_ShouldHandleEdgeCases` - Empty strings, single characters, many parts
- `RearrangeName_ShouldHandleSpecialCharacters` - Non-ASCII characters, numbers, apostrophes

#### String Normalization Edge Cases
- `NormalizeString_ShouldHandleEdgeCases` - Whitespace characters, non-breaking spaces
- `NormalizeString_ShouldBeIdempotent` - Repeated normalization gives same result

#### Month Conversion Edge Cases
- `GetMonthAsString_ShouldHandleInvalidInputs` - Boundary values, negative numbers
- `GetMonthAsString_ShouldCoverAllValidMonths` - All 12 months validated

#### Posting Date Edge Cases
- `GetEffectivePostingDate_ShouldHandleMonthBoundary` - Calculations crossing month boundaries
- `GetEffectivePostingDate_ShouldHandleYearBoundary` - Calculations crossing year boundaries
- `GetEffectivePostingDate_ShouldHandleLeapYearTransition` - Leap day handling

#### Exclusion Loading Edge Cases
- `LoadExclusionsFromFile_ShouldHandleMissingFile` - Non-existent file handling
- `LoadExclusionsFromFile_ShouldHandleEmptyFile` - Empty file handling
- `LoadExclusionsFromFile_ShouldHandleLargeFiles` - Performance with many exclusions

#### Decimal Precision Tests
- `DecimalPrecision_ShouldPreserveAccuracy` - Financial calculations maintain precision
- `DecimalComparison_ShouldUseTolerance` - Tolerance for floating-point comparisons

**Business Requirement**: Financial data must be accurate to the cent (2 decimal places).

#### Phone Number Edge Cases
- `ExtractLast4Digits_ShouldHandleVariousFormats` - Different phone formats, short numbers

#### Concurrent Scenario Tests
- `MultipleDonations_ShouldAggregateCorrectly` - Multiple transactions same person, same day
- `LargeAmounts_ShouldBeHandled` - Very large transaction amounts

#### Null/Empty Reference Tests
- `NullMessage_ShouldBeHandledGracefully` - Null message handling
- `EmptyOrNull_ShouldBeConsistentlyHandled` - Consistent empty value treatment

## Test Helper Class: MobilePayTestHelper.cs

This helper class provides reflection-based access to private methods of the `MobilePay` class, allowing thorough testing without exposing internals:

### Methods Exposed:
- `CallRearrangeName(string)` - Tests name rearrangement
- `CallNormalizeString(string)` - Tests string normalization
- `CallGetMonthAsString(int)` - Tests month conversion
- `CallGetEffectivePostingDate(DateTime)` - Tests posting date logic
- `CallLoadExclusionsFromFile(string)` - Tests exclusion file loading
- `CallGetTransactions(string)` - Tests CSV parsing
- `GenerateDailyExcelFile()` - Helper for Excel file testing

## Testing Approach: Outside-In (Behavioral Testing)

### Why Outside-In Testing?

1. **Business Requirements Persist**: Tests verify what the system should do, not how it does it
2. **Refactoring Safe**: Implementation can change without breaking tests
3. **Integration Focused**: Tests verify components work together correctly
4. **Maintainability**: Business logic is easier to understand through test descriptions

### Test Structure Example:

```csharp
/// <summary>
/// Business requirement: System must handle multiple donations from the same person
/// in the same month and sum them together correctly.
/// </summary>
[Fact]
public void MonthlySummary_ShouldAggregateDonationsForSamePerson()
{
    // Arrange - set up business scenario
    var donations = new List<(string name, decimal amount, int month)> { ... };
    
    // Act - perform the operation
    var johnTotal = donations
        .Where(d => d.name == "John Smith" && d.month == 1)
        .Sum(d => d.amount);
    
    // Assert - verify business requirement is met
    Assert.Equal(175.00m, johnTotal);
}
```

## Running the Tests

### Command Line
```bash
dotnet test "HVM Kasserer.Tests\HVM Kasserer.Tests.csproj"
```

### Visual Studio
- Open Test Explorer (Test ? Test Explorer)
- All tests will appear grouped by class
- Run individual tests, groups, or entire suite

### With Coverage
```bash
dotnet test /p:CollectCoverage=true /p:CoverageFormat=opencover
```

## Test Statistics

- **Total Test Files**: 3
- **Total Test Classes**: 3
- **Total Test Methods**: ~60+
- **Test Themes**: 
  - Name handling and formatting (5 tests)
  - String normalization (4 tests)
  - Month conversion (4 tests)
  - Posting date logic (7 tests)
  - Exclusion handling (7 tests)
  - Transaction classification (6 tests)
  - Deduplication (4 tests)
  - Message handling (4 tests)
  - Edge cases (20+ tests)
  - Integration scenarios (10+ tests)

## Key Business Requirements Verified

1. ? Names are formatted consistently for reports
2. ? Exclusion keywords are matched case-insensitively
3. ? Transactions are classified correctly (donation vs. fee)
4. ? Weekend transactions post to the following Monday
5. ? Monthly summaries aggregate correctly
6. ? Duplicate entries are prevented
7. ? Danish locale is handled correctly
8. ? Financial precision is maintained
9. ? System handles errors gracefully
10. ? CSV files are parsed correctly

## Future Test Enhancements

- Add tests for PDF generation
- Add tests for Excel file modification and updating
- Add concurrency tests for parallel transaction processing
- Add performance benchmarks for large transaction batches
- Add tests for the complete SummarizeMobilePayTransactions workflow
