using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace CvsReader
{
    class Program
    {
        static void Main(string[] args)
        {
            PrintReport();
            //List<Person> people = new List<Person>();

            //using (var reader = new StreamReader(@"C:\Work\GilecV4.csv"))
            //{
            //    while (!reader.EndOfStream)
            //    {
            //        var line = reader.ReadLine();
            //        var values = line.Split(';');
            //        people.Add(new Person
            //        {
            //            Name = values[1],
            //            Surname = values[0],
            //            Patronymic = values[2],
            //            Law = values[3],
            //            Category = values[4],
            //            StatementName = values[5],
            //            Pay = decimal.Parse(values[6], CultureInfo.InvariantCulture),
            //            DistrictName = int.Parse(values[7])
            //        });
            //    }
            //}
            //List<District> districtsNames = new List<District>();
            //string currentRajon = "";
            //using (var reader = new StreamReader(@"C:\Users\i.khakimzhanov\Desktop\rajon.csv"))
            //{

            //    while (!reader.EndOfStream)
            //    {
            //        var line = reader.ReadLine();
            //        var values = line.Split(';');
            //        districtsNames.Add(new District
            //        {
            //            kod_raj = int.Parse(values[0]),
            //            rajon = values[1]
            //        });
            //    }
            //}
            //var districts = people.Select(c => c.DistrictName).Distinct().ToList();
            //ExcelHelper excel = null;
            //foreach (var item in districts)
            //{
            //    excel = new ExcelHelper(@"C:\Users\i.khakimzhanov\Desktop\Report.xlsx", 1);
            //    try
            //    {
            //        List<Person> persons = people.Where(p => p.DistrictName == item).ToList();

            //        excel.WritePersons(persons);
            //        currentRajon = districtsNames.First(x => x.kod_raj == item).rajon;
            //        currentRajon = currentRajon.Replace('.', ' ');
            //        excel.WriteRajonHeaderForReport(currentRajon);
            //        excel.SetWidth();
            //        excel.SaveAs(currentRajon);
            //        Console.WriteLine(currentRajon + " DONE");

            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex.Message);
            //        excel.Close();
            //    }
            //    excel.Close();
            //}

        }

        static void PrintReport()
        {
            List<Payment> payments = new List<Payment>();

            decimal[] chaesLawArr = new decimal[] { 990000000008, 990000000568, 990000000106 };
            decimal maiakLaw = 990000000010;
            decimal semipalat = 990000000145;
            using (var reader = new StreamReader(@"C:\Work\V4.csv"))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');
                    payments.Add(new Payment(int.Parse(values[0]), decimal.Parse(values[1]), values[2], decimal.Parse(values[3], CultureInfo.InvariantCulture)));
                }
            }

            List<District> districtsNames = new List<District>();
            string currentRajon = "";
            using (var reader = new StreamReader(@"C:\Users\i.khakimzhanov\Desktop\rajon.csv"))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');
                    districtsNames.Add(new District
                    {
                        kod_raj = int.Parse(values[0]),
                        rajon = values[1]
                    });
                }
            }
            //var districts = payments.Select(c => c.kod_raj).Distinct().ToList();
            //foreach (var item in districts)
            //{
            ExcelHelper excel = new ExcelHelper(@"C:\Users\i.khakimzhanov\Desktop\Форма для заполнения_мин_мах_ср размер выплат.xlsx", 1);
            try
            {

                List<PaymentResult> chaesRes = payments.
                            Where(x =>
                            //x.kod_raj == item &&
                            chaesLawArr.Contains(x.nzp_law)).
                            GroupBy(x => x.str_name).
                            Select(l => new PaymentResult
                            {
                                Law = l.First().str_name.ToString(),
                                MimSumVipl = l.Min(c => c.sum_vipl),
                                MaxSumVipl = l.Max(c => c.sum_vipl),
                                AverageSumVipl = Math.Round(l.Average(c => c.sum_vipl), 2, MidpointRounding.AwayFromZero)
                            }).ToList();

                List<PaymentResult> maiakRes = payments.
                           Where(x =>
                           //x.kod_raj == item &&
                           x.nzp_law == maiakLaw).
                           GroupBy(x => x.str_name).
                           Select(l => new PaymentResult
                           {
                               Law = l.First().str_name.ToString(),
                               MimSumVipl = l.Min(c => c.sum_vipl),
                               MaxSumVipl = l.Max(c => c.sum_vipl),
                               AverageSumVipl = Math.Round(l.Average(c => c.sum_vipl), 2, MidpointRounding.AwayFromZero)
                           }).ToList();

                List<PaymentResult> semipalatRes = payments.
                           Where(x =>
                           //x.kod_raj == item &&
                           x.nzp_law == semipalat).
                           GroupBy(x => x.str_name).
                           Select(l => new PaymentResult
                           {
                               Law = l.First().str_name.ToString(),
                               MimSumVipl = l.Min(c => c.sum_vipl),
                               MaxSumVipl = l.Max(c => c.sum_vipl),
                               AverageSumVipl = Math.Round(l.Average(c => c.sum_vipl), 2, MidpointRounding.AwayFromZero)
                           }).ToList();

                foreach (var payment in chaesRes)
                {
                    int row = excel.FindRow(payment);
                    if (row != -1)
                        excel.Write(payment, row, LawEnum.Chaes);
                    else
                        Console.WriteLine("Problem: " + payment.Law);
                }
                foreach (var payment in maiakRes)
                {
                    int row = excel.FindRow(payment);
                    if (row != -1)
                        excel.Write(payment, row, LawEnum.Maiak);
                    else
                        Console.WriteLine("Problem: " + payment.Law);

                }
                foreach (var payment in semipalatRes)
                {
                    int row = excel.FindRow(payment);
                    if (row != -1)
                        excel.Write(payment, row, LawEnum.Semipalat);
                    else
                        Console.WriteLine("Problem: " + payment.Law);
                }
                //currentRajon = districtsNames.First(x => x.kod_raj == item).rajon;
                //currentRajon = currentRajon.Replace('.', ' ');
                //excel.WriteRajonHeader(districtsNames.First(x => x.kod_raj == item).rajon);
                excel.WriteRajonHeader("Республика Татарстан");
                excel.SetZeroValue();
                excel.SaveAs("Республика Татарстан");
                Console.WriteLine("Республика Татарстан" + " DONE");
            }
            finally
            {
                excel.Close();
            }
            //}
        }
    }
}
