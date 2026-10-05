namespace BWG_WasteScraper
{
    internal record AcceptanceRow(
        string Id,
        string Sequence,
        string RequestNumber,
        string FactoryName,
        string FactoryRegNo,
        string SubmittedAt,
        string Status,
        string RequestType,
        string Deadline);
}