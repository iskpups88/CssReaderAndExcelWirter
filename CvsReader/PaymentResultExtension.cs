namespace CvsReader
{
    static class PaymentResultExtension
    {
        public static string[] ToArrayPaymentResult(this PaymentResult result)
        {
            string[] arr = new string[] { result.MimSumVipl.ToString(), result.MaxSumVipl.ToString(), result.AverageSumVipl.ToString() };
            return arr;
        }

        public static string[] ToArrayPerson(this Person result)
        {
            string[] arr = new string[] { result.Surname, result.Name, result.Patronymic, result.Law, result.StatementName, result.Category, result.Pay.ToString() };
            return arr;
        }
    }
}
