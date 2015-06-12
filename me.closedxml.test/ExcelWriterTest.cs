using System.Collections.ObjectModel;
using NUnit.Framework;

namespace closedxml.test
{
    public class ExcelWriterTest
    {
        [Test]
        public void WriteTest()
        {
            var companies = new Collection<Company>();
            companies.Add(new Company() { CompanyId = 1, Name = "TAB Consultores" });
            companies.Add(new Company() { CompanyId = 2, Name = "Async Consultores" });

            var customers = new Collection<Customer>();
            customers.Add(new Customer() { CustomerId = 1, Name = "Miguel Ángel Martín Hernández" });
            customers.Add(new Customer() { CustomerId = 2, Name = "Montserrar Gómez Rubiano" });
            customers.Add(new Customer() { CustomerId = 3, Name = "Miguel Martín Sánchez" });
            customers.Add(new Customer() { CustomerId = 4, Name = "María Francisca Hernández Jiménez" });

            var data = new Collection<ExcelData> 
            { 
                new ExcelData() { Name = "Companies", Data = companies}, 
                new ExcelData() { Name = "Customers", Data= customers }
            };
            var excelWriter = new ExcelWriter(data);
            excelWriter.Write();
        }
    }
}