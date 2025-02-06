namespace FuzzySheets.Domain.Config
{
    public class MutationConfig
    {
        public MutationDetail NumberFormat { get; set; } = new MutationDetail();
        public MutationDetail DateFormat { get; set; } = new MutationDetail();
        public MutationDetail Encoding { get; set; } = new MutationDetail();
        public MutationDetail EmptyValues { get; set; } = new MutationDetail();
        public MutationDetail DecimalPrecision { get; set; } = new MutationDetail();
    }
}