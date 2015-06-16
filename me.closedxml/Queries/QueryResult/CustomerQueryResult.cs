using System;

namespace me.closedxml.Queries.QueryResult
{
    public class CustomerQueryResult : IQueryResult
    {
        public CustomerQueryResult(int customerId, string name, DateTime birthDate, decimal lastInvoice, bool removed)
        {
            CustomerId = customerId;
            Name = name;
            BirthDate = birthDate;
            LastInvoice = lastInvoice;
            Removed = removed;
        }
        public int CustomerId { get; set; }
        public string Name { get; set; }
        public DateTime BirthDate { get; set; }
        public decimal LastInvoice { get; set; }
        public bool Removed { get; set; }
    }
}