using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVM_Kasserer
{
    class HelperMethods
    {
        public static string ExtractLast4Digits(string phoneNumber)
        {
            // Remove spaces and non-numeric characters
            var digits = new string(phoneNumber.Where(char.IsDigit).ToArray());

            // Return the last 4 digits if available
            return digits.Length >= 4 ? digits[^4..] : string.Empty;
        }
    }
}
