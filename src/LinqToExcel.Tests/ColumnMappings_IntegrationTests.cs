using System;
using System.Linq;
using MbUnit.Framework;
using System.IO;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [FixtureCategory("Integration")]
    [TestFixture]
    public class ColumnMappingsIntegrationTests 
    {
        ExcelQueryFactory repo;
        string excelFileName;
        string worksheetName;

        [TestFixtureSetUp]
        public void Fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            this.excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
            this.worksheetName = "ColumnMappings";
        }

        [SetUp]
        public void S()
        {
            this.repo = new ExcelQueryFactory(new LogManagerFactory());
            this.repo.FileName = this.excelFileName;
        }

        [Test]
        public void all_properties_have_column_mappings()
        {
            this.repo.AddMapping<Company>(x => x.Name, "Company Title");
            this.repo.AddMapping<Company>(x => x.CEO, "Boss");
            this.repo.AddMapping<Company>(x => x.EmployeeCount, "Number of People");
            this.repo.AddMapping<Company>(x => x.StartDate, "Initiation Date");

            var companies = from c in this.repo.Worksheet<Company>(this.worksheetName)
                            where c.Name == "Taylor University"
                            select c;

            var rival = companies.ToList().First();
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Taylor University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(400, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
        }

        [Test]
        public void some_properties_have_column_mappings()
        {
            this.repo.AddMapping<Company>(x => x.CEO, "Boss");
            this.repo.AddMapping<Company>(x => x.StartDate, "Initiation Date");

            var companies = from c in this.repo.Worksheet<Company>(this.worksheetName)
                            where c.Name == "Anderson University"
                            select c;

            Company rival = companies.ToList()[0];
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Anderson University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(300, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
        }


        [Test]
        public void column_mappings_with_transformation()
        {
            this.repo.AddMapping<Company>(x => x.IsActive, "Active", x => x == "Y");
            var companies = from c in this.repo.Worksheet<Company>(this.worksheetName)
                            select c;

            foreach (var company in companies)
                Assert.AreEqual(company.StartDate > new DateTime(1980, 1, 1), company.IsActive);
        }

        [Test]
        public void Transformation()
        {
            //Add transformation to change the Name value to 'Looney Tunes' if it is originally 'ACME'
            this.repo.AddTransformation<Company>(p => p.Name, value => (value == "ACME") ? "Looney Tunes" : value);
            var firstCompany = (from c in this.repo.Worksheet<Company>(this.worksheetName)
                                select c).First();

            Assert.AreEqual("Looney Tunes", firstCompany.Name);
        }

        [Test]
        public void transformation_that_returns_null()
        {
            //Add transformation to change the Name value to 'Looney Tunes' if it is originally 'ACME'
            this.repo.AddTransformation<Company>(p => p.Name, value => null);
            var firstCompany = (from c in this.repo.Worksheet<Company>(this.worksheetName)
                                select c).First();

            Assert.AreEqual(null, firstCompany.Name);
        }

        [Test]
        public void annotated_properties_map_to_columns()
        {
            var companies = from c in this.repo.Worksheet<CompanyWithColumnAnnotations>(this.worksheetName)
                            where c.Name == "Taylor University"
                            select c;

            var rival = companies.ToList().First();
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Taylor University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(400, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
            Assert.AreEqual("N", rival.IsActive, "IsActive");
        }
    }
}
