using Xunit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HVM_Kasserer_Tests
{
    /// <summary>
    /// Tests for how transactions are sorted into donations and excluded (food, kontingent, events, etc.).
    /// Scenarios are modelled on real days from the MobilePay reports.
    /// </summary>
    public class MobilePayExclusionRuleTests : IDisposable
    {
        private readonly string _testDataDirectory;
        private readonly MobilePayTestHelper _helper = new();

        private static readonly List<string> Exclusions = new()
        {
            "mad", "frokost", "middag", "menighedsdag", "menigheds dag", "menighedsfag", "koncert", "entre", "måltid"
        };

        public MobilePayExclusionRuleTests()
        {
            _testDataDirectory = Path.Combine(Path.GetTempPath(), "MobilePayExclusionTests", Guid.NewGuid().ToString());
            Directory.CreateDirectory(_testDataDirectory);
        }

        #region Fixtures

        private record Tx(string Id, string Time, decimal Amount, string Message, string Name);

        /// <summary>
        /// Writes a CSV in the MobilePay report format with a Betaling and a Gebyr row per transaction.
        /// </summary>
        private string CreateCsv(string date, params Tx[] transactions)
        {
            var lines = new List<string> { "CVR Nummer;Virksomhedens Navn;Produkt;Betalingssted ID;Betalingsted Navn;Type;Beløb;Valuta;Tid;Dato;Tidsstempel;Besked;Reference;Transaktionsmærkat;Transaktions ID;Kundenavn;Mobilnummer" };
            foreach (var t in transactions)
            {
                var timestamp = $"{date}T{t.Time}:00+02:00";
                lines.Add($"1;HV;P;ID;HV;Betaling;{t.Amount:0},00;DKK;{t.Time};{date};{timestamp};{t.Message};R;;{t.Id};{t.Name};xxxx 1234");
                lines.Add($"1;HV;P;ID;HV;Gebyr;-1,00;DKK;{t.Time};{date};{timestamp};;R;;{t.Id};{t.Name};xxxx 1234");
            }

            var path = Path.Combine(_testDataDirectory, $"{Guid.NewGuid()}.csv");
            File.WriteAllLines(path, lines);
            return path;
        }

        private HashSet<string> Classify(string csvPath) =>
            _helper.CallGetExcludedTransactionIDs(csvPath, Exclusions);

        #endregion

        /// <summary>
        /// Menighedsdag 2026-09-20: collection for SM and payments for food came in at the same time.
        /// Only payments explicitly marked with the collection purpose count as donations.
        /// </summary>
        [Fact]
        public void Menighedsdag_OnlyExplicitlyMarkedPaymentsAreDonations()
        {
            var csv = CreateCsv("2026-09-20",
                new Tx("soldatermissionen", "11:04", 150, "Soldatermissionen", "Bodil Bådebygger"),
                new Tx("sm1", "11:43", 100, "SM", "Anne-Mette Ananasspiser"),
                new Tx("smgave", "11:44", 500, "SM GAVE", "Bjarne Bjergbestiger"),
                new Tx("sm2", "11:45", 250, "SM", "Jette Jonglør"),
                new Tx("kajkirsten", "11:47", 200, "\" Kaj og Kirsten \"", "Kirsten Kanopadler"),
                new Tx("valgmenighedsdag", "11:47", 250, "Valgmenighedsdag", "Svend Svampesamler"),
                new Tx("frokost", "11:50", 200, "Lone og Bodil Frokost", "Bodil Bådebygger"),
                new Tx("blank1", "11:53", 200, "", "Knud Kartoffelskræller"),
                new Tx("sm3", "12:00", 700, "SM", "Henning Havenisse"),
                new Tx("menighedsdag", "13:48", 100, "Menighedsdag", "Jørgen Juletræ"),
                new Tx("blank2", "14:05", 250, "", "Mette Marcipan"));

            var excluded = Classify(csv);

            Assert.Equal(
                new[] { "blank1", "blank2", "frokost", "kajkirsten", "menighedsdag", "valgmenighedsdag" },
                excluded.OrderBy(x => x));
        }

        /// <summary>
        /// On a Sunday with food after the start date, payments without a message close in time to
        /// the food payments are excluded, while explicitly marked donations are kept.
        /// </summary>
        [Fact]
        public void FoodPayments_ExcludeUnmarkedPaymentsWithinTimeWindow()
        {
            var csv = CreateCsv("2026-10-04",
                new Tx("dfs1", "10:30", 200, "DFS", "Jørgen Juletræ"),
                new Tx("dfs2", "11:35", 100, "Dfs", "Søren Sofakigger"),
                new Tx("mad1", "11:40", 100, "Mad", "Gerda Gummistøvle"),
                new Tx("blank", "11:42", 100, "", "Rikke Rulleskøjte"),
                new Tx("personal", "11:45", 100, "Rikke og Kalle", "Kalle Kaffekande"),
                new Tx("gave", "11:50", 100, "Gave", "Tina Tennisspiller"),
                new Tx("ordet1", "11:51", 100, "Ordet og Israel", "Lise Lakridsrod"),
                new Tx("ordet2", "11:52", 100, "Ordet & Israel", "Bodil Bådebygger"),
                new Tx("single", "11:53", 150, "Soldatermissionen", "Kurt Kolbøtte"),
                new Tx("mad2", "12:10", 100, "2xMad", "Niels Nøddeknækker"),
                new Tx("blankLater", "12:39", 50, "", "Rasmus Regnorm"),
                new Tx("blankMuchLater", "12:41", 50, "", "Mogens Mudderpøl"));

            var excluded = Classify(csv);

            // "single" is an organisation name written by only one payer close to the food payments,
            // so it is not recognised as a collection purpose
            Assert.Equal(
                new[] { "blank", "blankLater", "mad1", "mad2", "personal", "single" },
                excluded.OrderBy(x => x));
        }

        /// <summary>
        /// Church at 9 with donations and food at 13: the donations in the morning are too far
        /// from the food payments to be affected.
        /// </summary>
        [Fact]
        public void FoodPaymentsLaterInTheDay_DoNotAffectMorningDonations()
        {
            var csv = CreateCsv("2026-10-11",
                new Tx("blankChurch1", "09:45", 200, "", "Birthe Blomkål"),
                new Tx("blankChurch2", "10:15", 100, "", "Kurt Kolbøtte"),
                new Tx("personalChurch", "10:20", 100, "Hugo.", "Hugo Hængekøje"),
                new Tx("mad", "13:00", 100, "Mad", "Gerda Gummistøvle"),
                new Tx("blankFood", "13:02", 100, "", "Rikke Rulleskøjte"));

            var excluded = Classify(csv);

            Assert.Equal(new[] { "blankFood", "mad" }, excluded.OrderBy(x => x));
        }

        /// <summary>
        /// Regression, DFS 2026-02-15: days before the strict rule's start date are already booked
        /// and must be sorted as before, so payments without a message are still donations.
        /// </summary>
        [Fact]
        public void BeforeStartDate_FrokostDoesNotAffectPaymentsWithoutMessage()
        {
            var csv = CreateCsv("2026-02-15",
                new Tx("dfs1", "10:30", 200, "DFS", "Jørgen Juletræ"),
                new Tx("dfs2", "10:31", 50, "DF’s", "Søren Sofakigger"),
                new Tx("blank1", "11:40", 150, "", "Bodil Bådebygger"),
                new Tx("frokost1", "11:41", 50, "Frokost", "Gerda Gummistøvle"),
                new Tx("blank2", "11:42", 50, "", "Rikke Rulleskøjte"),
                new Tx("frokost2", "11:50", 100, "2 x frokost", "Niels Nøddeknækker"));

            var excluded = Classify(csv);

            Assert.Equal(new[] { "frokost1", "frokost2" }, excluded.OrderBy(x => x));
        }

        /// <summary>
        /// Regression, Ordet og Israel 2026-04-02: a meal evening before the start date keeps
        /// personal messages and payments without message as donations.
        /// </summary>
        [Fact]
        public void BeforeStartDate_MealEveningKeepsNonExplicitDonations()
        {
            var csv = CreateCsv("2026-04-02",
                new Tx("paaske", "18:07", 100, "Påskemåltid", "Jørgen Juletræ"),
                new Tx("blank", "18:08", 100, "", "Tina Tennisspiller"),
                new Tx("gave", "19:36", 75, "Gave Ordet og Israel", "Lise Lakridsrod"),
                new Tx("maaltid", "19:36", 200, "Måltid", "Gerda Gummistøvle"),
                new Tx("personal", "19:37", 100, "Gave", "Tina Tennisspiller"));

            var excluded = Classify(csv);

            Assert.Equal(new[] { "maaltid", "paaske" }, excluded.OrderBy(x => x));
        }

        /// <summary>
        /// The strict rule only applies on the menighedsdag itself, not on other days in the same report.
        /// </summary>
        [Fact]
        public void Menighedsdag_DoesNotAffectOtherDays()
        {
            var menighedsdag = File.ReadAllLines(CreateCsv("2026-09-20",
                new Tx("menighedsdag", "12:00", 200, "Menighedsdag", "Sofie Sæbeboble"),
                new Tx("blankSameDay", "12:01", 200, "", "Hugo Hængekøje")));
            var nextSunday = File.ReadAllLines(CreateCsv("2026-09-27",
                new Tx("puls", "10:30", 200, "PULS", "Birthe Blomkål"),
                new Tx("blankOtherDay", "10:31", 100, "", "Kurt Kolbøtte")));

            var csv = Path.Combine(_testDataDirectory, "combined.csv");
            File.WriteAllLines(csv, menighedsdag.Concat(nextSunday.Skip(1)));

            var excluded = Classify(csv);

            Assert.Equal(new[] { "blankSameDay", "menighedsdag" }, excluded.OrderBy(x => x));
        }

        /// <summary>
        /// A collection purpose written by at least two payers counts as explicit, even if it
        /// is not in the collection purposes overview, and regardless of case and punctuation.
        /// </summary>
        [Fact]
        public void Menighedsdag_RecurringMessageIsTreatedAsCollectionPurpose()
        {
            var csv = CreateCsv("2026-09-27",
                new Tx("puls1", "11:40", 100, "Puls", "Jørgen Juletræ"),
                new Tx("puls2", "11:52", 100, "PULS.", "Birthe Blomkål"),
                new Tx("menighedsdag", "11:48", 200, "Menigheds dag.", "Hugo Hængekøje"),
                new Tx("single", "11:53", 200, "Tak for i dag", "Henning Havenisse"));

            var excluded = Classify(csv);

            Assert.Equal(new[] { "menighedsdag", "single" }, excluded.OrderBy(x => x));
        }

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
    }
}
