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
    public class ExcelWriterReaderTest
    {
        private const string FilePath = @"c:\temp\excel.xlsx";
        private Collection<IExcelData<IQueryResult>> _originalItems;
        private Random _random;
        [TestFixtureSetUp]
        public void TextFixtureSetUp()
        {
            _random = new Random();
            _originalItems = new Collection<IExcelData<IQueryResult>>();
            _originalItems.Add(new ExcelData<CompanyExcelConfigurationWorksheetRow>("Companies", GetCompanies()));
            _originalItems.Add(new ExcelData<CustomerExcelConfigurationWorksheetRow>("Customers", GetCustomers()));
        }

        [Test]
        public void ExcelWriterTest()
        {
            var excelWriter = new ExcelWriter(FilePath, _originalItems);
            excelWriter.Write();
        }

        [Test]
        public void ExcelReaderTest()
        {
            var excelReader = new ExcelReader(FilePath);
            var items = excelReader.Read().ToList();

            Assert.AreEqual(_originalItems.Count, items.Count());
            foreach (var item in items)
            {
                var worksheetName = item.WorksheetName;
                var data = _originalItems.Single(p => p.WorksheetName == worksheetName);
                Assert.AreEqual(data.Data.Count(), item.Data.Count());
            }
        }

        #region Helpers

        private IEnumerable<CompanyQueryResult> GetCompanies()
        {
            const int companiesNumber = 50;
            var companies = new Collection<CompanyQueryResult>();
            for (var i = 0; i < companiesNumber; i++)
            {
                var companyQueryResult = new CompanyQueryResult(i, Guid.NewGuid().ToString());
                companies.Add(companyQueryResult);
            }
            return companies;
        }

        private IEnumerable<CustomerQueryResult> GetCustomers()
        {


            var customers = new Collection<CustomerQueryResult>();
            customers.Add(new CustomerQueryResult(1, "Miguel Ángel Martín Hernández", DateTime.UtcNow.Date, (decimal)_random.NextDouble(), true));
            customers.Add(new CustomerQueryResult(2, "Montserrar Gómez Rubiano", DateTime.UtcNow.AddYears(-10).Date, (decimal)_random.NextDouble(), false));
            customers.Add(new CustomerQueryResult(3, "Miguel Martín Sánchez", DateTime.UtcNow.AddYears(-20).Date, (decimal)_random.NextDouble(), true));
            customers.Add(new CustomerQueryResult(4, "María Francisca Hernández Jiménez", DateTime.UtcNow.AddYears(-30).Date, (decimal)_random.NextDouble(), false));
            return customers;
        }

        #endregion
    }
}