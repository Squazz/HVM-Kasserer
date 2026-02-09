using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace HVM_Kasserer_Tests
{
    /// <summary>
    /// Helper class to access private methods of MobilePay for testing purposes.
    /// Uses reflection to invoke private methods in a type-safe manner.
    /// </summary>
    internal class MobilePayTestHelper
    {
        private static readonly Type MobilePayType;
        private static readonly Type HelperMethodsType;
        
        private static readonly MethodInfo RearrangeNameMethod;
        private static readonly MethodInfo NormalizeStringMethod;
        private static readonly MethodInfo GetMonthAsStringMethod;
        private static readonly MethodInfo GetEffectivePostingDateMethod;
        private static readonly MethodInfo LoadExclusionsFromFileMethod;
        private static readonly MethodInfo GetTransactionsMethod;
        private static readonly MethodInfo ExtractLast4DigitsMethod;

        private object? _mobilePayInstance;

        static MobilePayTestHelper()
        {
            // Load types using reflection to avoid visibility issues
            var assemblyName = "HVM Kasserer";
            var assembly = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(a => a.GetName().Name == assemblyName);
            
            if (assembly == null)
            {
                assembly = System.Reflection.Assembly.Load(assemblyName);
            }

            MobilePayType = assembly.GetType("HVM_Kasserer.MobilePay") 
                ?? throw new TypeLoadException("Could not load MobilePay type");
            
            HelperMethodsType = assembly.GetType("HVM_Kasserer.HelperMethods")
                ?? throw new TypeLoadException("Could not load HelperMethods type");

            RearrangeNameMethod = MobilePayType.GetMethod(
                "RearrangeName", BindingFlags.NonPublic | BindingFlags.Instance);
            
            NormalizeStringMethod = MobilePayType.GetMethod(
                "NormalizeString", BindingFlags.NonPublic | BindingFlags.Instance);
            
            GetMonthAsStringMethod = MobilePayType.GetMethod(
                "GetMonthAsString", BindingFlags.NonPublic | BindingFlags.Instance);
            
            GetEffectivePostingDateMethod = MobilePayType.GetMethod(
                "GetEffectivePostingDate", BindingFlags.NonPublic | BindingFlags.Instance);
            
            LoadExclusionsFromFileMethod = MobilePayType.GetMethod(
                "LoadExclusionsFromFile", BindingFlags.NonPublic | BindingFlags.Instance);
            
            GetTransactionsMethod = MobilePayType.GetMethod(
                "GetTransactions", BindingFlags.NonPublic | BindingFlags.Instance);
            
            ExtractLast4DigitsMethod = HelperMethodsType.GetMethod(
                "ExtractLast4Digits", BindingFlags.Public | BindingFlags.Static);
        }

        public MobilePayTestHelper()
        {
            try
            {
                _mobilePayInstance = Activator.CreateInstance(MobilePayType);
            }
            catch
            {
                // Instance creation might fail due to file paths
                _mobilePayInstance = null;
            }
        }

        /// <summary>
        /// Calls the private RearrangeName method on MobilePay.
        /// </summary>
        public string CallRearrangeName(string fullName)
        {
            var instance = GetOrCreateInstance();
            var result = RearrangeNameMethod?.Invoke(instance, new object[] { fullName });
            return result?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Calls the private NormalizeString method on MobilePay.
        /// </summary>
        public string CallNormalizeString(string input)
        {
            var instance = GetOrCreateInstance();
            var result = NormalizeStringMethod?.Invoke(instance, new object[] { input });
            return result?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Calls the private GetMonthAsString method on MobilePay.
        /// </summary>
        public string? CallGetMonthAsString(int monthNumber)
        {
            var instance = GetOrCreateInstance();
            var result = GetMonthAsStringMethod?.Invoke(instance, new object[] { monthNumber });
            return result?.ToString();
        }

        /// <summary>
        /// Calls the private GetEffectivePostingDate method on MobilePay.
        /// </summary>
        public DateTime CallGetEffectivePostingDate(DateTime dateTime)
        {
            var instance = GetOrCreateInstance();
            var result = GetEffectivePostingDateMethod?.Invoke(instance, new object[] { dateTime });
            return result is DateTime dt ? dt : dateTime;
        }

        /// <summary>
        /// Calls the private LoadExclusionsFromFile method on MobilePay.
        /// </summary>
        public List<string> CallLoadExclusionsFromFile(string filePath)
        {
            var instance = GetOrCreateInstance();
            
            // Set the exclusionsFilePath field
            var exclusionsFilePathField = MobilePayType.GetField("exclusionsFilePath", BindingFlags.NonPublic | BindingFlags.Instance);
            if (exclusionsFilePathField != null)
            {
                exclusionsFilePathField.SetValue(instance, filePath);
            }

            var result = LoadExclusionsFromFileMethod?.Invoke(instance, null);
            return result as List<string> ?? new List<string>();
        }

        /// <summary>
        /// Calls the private GetTransactions method on MobilePay.
        /// </summary>
        public dynamic CallGetTransactions(string csvFilePath)
        {
            var instance = GetOrCreateInstance();
            
            // Set the mobilePayFilepath field
            var mobilePayFilepathField = MobilePayType.GetField("mobilePayFilepath", BindingFlags.NonPublic | BindingFlags.Instance);
            if (mobilePayFilepathField != null)
            {
                mobilePayFilepathField.SetValue(instance, csvFilePath);
            }

            var result = GetTransactionsMethod?.Invoke(instance, null);
            return result;
        }

        /// <summary>
        /// Calls the public static ExtractLast4Digits method from HelperMethods.
        /// </summary>
        public string CallExtractLast4Digits(string phoneNumber)
        {
            var result = ExtractLast4DigitsMethod?.Invoke(null, new object[] { phoneNumber });
            return result?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Helper method to generate a daily Excel file for testing.
        /// </summary>
        public string GenerateDailyExcelFile(DateTime date, List<(string name, decimal amount, string type)> transactionData, string outputPath)
        {
            string fileName = $"Mobilepay-{date:yyyy-MM-dd}.xlsx";
            string filePath = Path.Combine(outputPath, fileName);

            using var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Transactions");

            // Headers
            var headers = new[] { "Date", "Time", "Name", "Phone", "Type", "Amount", "Message", "TransactionID", "Donation" };
            for (int i = 0; i < headers.Length; i++)
                ws.Cell(1, i + 1).Value = headers[i];

            int row = 2;
            foreach (var (name, amount, type) in transactionData)
            {
                ws.Cell(row, 1).Value = date.ToString("yyyy-MM-dd");
                ws.Cell(row, 2).Value = DateTime.Now.ToString("HH:mm:ss");
                ws.Cell(row, 3).Value = name;
                ws.Cell(row, 4).Value = "1234";
                ws.Cell(row, 5).Value = type;
                ws.Cell(row, 6).Value = amount;
                ws.Cell(row, 7).Value = string.Empty;
                ws.Cell(row, 8).Value = $"TXN{row:000}";
                ws.Cell(row, 9).Value = "Yes";
                row++;
            }

            // Sum row
            ws.Cell(row, 5).Value = "Total";
            ws.Cell(row, 6).FormulaA1 = $"=SUM(F2:F{row - 1})";

            ws.Column(6).Style.NumberFormat.Format = "#,##0.00";
            ws.Columns().AdjustToContents();

            workbook.SaveAs(filePath);
            return fileName;
        }

        /// <summary>
        /// Gets or creates a MobilePay instance for testing.
        /// </summary>
        private object GetOrCreateInstance()
        {
            _mobilePayInstance ??= Activator.CreateInstance(MobilePayType);
            return _mobilePayInstance ?? throw new InvalidOperationException("Could not create MobilePay instance");
        }
    }
}
