namespace me.closedxml.Queries.QueryResult
{
    public class CompanyQueryResult : IQueryResult
    {
        public CompanyQueryResult(int companyId, string name)
        {
            CompanyId = companyId;
            Name = name;
        }

        public int CompanyId { get; set; }
        public string Name { get; set; }
    }
}