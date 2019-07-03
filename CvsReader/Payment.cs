namespace CvsReader
{
    class Payment
    {
        public int kod_raj { get; set; }
        public decimal nzp_law { get; set; }
        public string str_name { get; set; }
        public decimal sum_vipl { get; set; }


        public Payment(int kod_raj, decimal nzp_law, string srt_name, decimal sum_vipl)
        {
            this.kod_raj = kod_raj;
            this.nzp_law = nzp_law;
            this.sum_vipl = sum_vipl;
            this.str_name = srt_name;
        }
    }
}
