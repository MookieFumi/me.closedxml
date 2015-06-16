using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ClosedXML.Excel;
using me.closedxml;
using me.closedxml.Queries.QueryResult;
using me.closedxml.Reader;
using me.closedxml.Writer;
using NUnit.Framework;

namespace me.closedxml.test
{
    public class ExcelTest
    {
        private const string FilePath = @"c:\temp\excel.xlsx";
        private Collection<IExcelData<IQueryResult>> _items;

        [TestFixtureSetUp]
        public void TextFixtureSetUp()
        {
            _items = new Collection<IExcelData<IQueryResult>>();
            _items.Add(new ExcelData<CompanyExcelConfigurationWorksheetRow>("Companies", GetCompanies()));
            _items.Add(new ExcelData<CustomerExcelConfigurationWorksheetRow>("Customers", GetCustomers()));
        }

        [Test]
        public void WriteTest()
        {
            var excelWriter = new ExcelWriter(FilePath, _items);
            excelWriter.Write();
        }

        [Test]
        public void ReadTest()
        {
            var excelReader = new ExcelReader(FilePath);
            var itemsRead = excelReader.Read().ToList();

            Assert.AreEqual(_items.Count, itemsRead.Count());
            foreach (var itemRead in itemsRead)
            {
                var worksheetName = itemRead.WorksheetName;
                var data = _items.Single(p => p.WorksheetName == worksheetName);

                Assert.AreEqual(data.Data.Count(), itemRead.Data.Count());
            }
        }

        #region Helpers

        private IEnumerable<CompanyQueryResult> GetCompanies()
        {
            var companies = new Collection<CompanyQueryResult>
                            {
                                new CompanyQueryResult(1, "TAB Consultores"),
                                new CompanyQueryResult(2, "Async Consultores")
                            };
            return companies;
        }

        private IEnumerable<CustomerQueryResult> GetCustomers()
        {
            var random = new Random();

            var customers = new Collection<CustomerQueryResult>();
            customers.Add(new CustomerQueryResult(1, "Miguel Ángel Martín Hernández", DateTime.UtcNow.Date, (decimal)random.NextDouble(), true));
            customers.Add(new CustomerQueryResult(2, "Montserrar Gómez Rubiano", DateTime.UtcNow.AddYears(-10).Date, (decimal)random.NextDouble(), false));
            customers.Add(new CustomerQueryResult(3, "Miguel Martín Sánchez", DateTime.UtcNow.AddYears(-20).Date, (decimal)random.NextDouble(), true));
            customers.Add(new CustomerQueryResult(4, "María Francisca Hernández Jiménez", DateTime.UtcNow.AddYears(-30).Date, (decimal)random.NextDouble(), false));
            return customers;
        }

        #endregion
    }
}