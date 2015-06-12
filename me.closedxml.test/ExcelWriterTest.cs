using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using me.closedxml.Queries.QueryResult;
using NUnit.Framework;

namespace me.closedxml.test
{
    public class ExcelWriterTest
    {
        [Test]
        public void WriteTest()
        {
            var data = new Collection<ExcelData> 
            { 
                new ExcelData { Name = "Companies", Data = GetCompanies()}, 
                new ExcelData { Name = "Customers", Data= GetCustomers ()}
            };
            var excelWriter = new ExcelWriter(data);
            excelWriter.Write();
        }

        #region Helpers

        private IEnumerable<CompanyQueryResult> GetCompanies()
        {
            var companies = new Collection<CompanyQueryResult>();
            companies.Add(new CompanyQueryResult { CompanyId = 1, Name = "TAB Consultores" });
            companies.Add(new CompanyQueryResult { CompanyId = 2, Name = "Async Consultores" });
            return companies;
        }

        private IEnumerable<CustomerQueryResult> GetCustomers()
        {
            var customers = new Collection<CustomerQueryResult>();
            customers.Add(new CustomerQueryResult { CustomerId = 1, Name = "Miguel Ángel Martín Hernández" });
            customers.Add(new CustomerQueryResult { CustomerId = 2, Name = "Montserrar Gómez Rubiano" });
            customers.Add(new CustomerQueryResult { CustomerId = 3, Name = "Miguel Martín Sánchez" });
            customers.Add(new CustomerQueryResult { CustomerId = 4, Name = "María Francisca Hernández Jiménez" });
            return customers;
        }

        #endregion
    }
}