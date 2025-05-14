using HVM_Kasserer;

class Program
{
    static void Main()
    {
        var bankPosteringer = new BankPosteringer();
        bankPosteringer.HandleBankPosteringer();

        var mobilePay = new MobilePay();
        mobilePay.SummarizeMobilePayTransactions();
    }
}
