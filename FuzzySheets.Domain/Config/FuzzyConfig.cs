namespace FuzzySheets.Domain.Config
{
    public class FuzzyConfig
    {
        public MutationConfig MutationConfig { get; set; } = new MutationConfig();
        public BatchProcessing BatchProcessing { get; set; } = new BatchProcessing();
    }
}