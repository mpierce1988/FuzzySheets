namespace FuzzySheets.Domain.Config
{
    public class MutationDetail
    {
        public bool IsEnabled { get; set; }
        public double Strength { get; set; }
        public double PercentMutated { get; set; }
        public bool IsNumeric { get; set; }
        public bool IsDate { get; set; }
        public bool IsString { get; set; }
    }
}