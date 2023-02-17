using ExcelAnalysis.Models;
using LinqToExcel;

namespace ExcelAnalysis.Helper
{
    public static class ExtractDuplicates
    {
        public static void Run()
        {
            string actualPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            var _excelFile = new ExcelQueryFactory(actualPath + "\\Excels\\countries.xlsx");

            //open excel file using linq to excel
            var file = _excelFile.GetWorksheetNames();

            //open spreadsheet
            var listCountries = from c in _excelFile.Worksheet<CountriesInDb>("country_202302162148")
                select c;

            var haveTheSameCode = listCountries.ToList().GroupBy(c => c.code).Where(g => g.Count() > 1).ToList();
            foreach (var countriesGrouped in haveTheSameCode)
            {
                Console.WriteLine(string.Join(" ", countriesGrouped.Select(s => s.country_name)));
            }

            //Write 3 console lines with a wide liner
            
            Console.WriteLine("--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("--------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("--------------------------------------------------------------------------------------------------------------------");
        }

        //return an list of countries in db as List<>
        public static List<CountriesInDb> GetCountriesInDb()
        {
            string actualPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var _excelFile = new ExcelQueryFactory(actualPath + "\\Excels\\countries.xlsx");

            //open excel file using linq to excel
            //var file = _excelFile.GetWorksheetNames();

            //open spreadsheet
            var listCountries = from c in _excelFile.Worksheet<CountriesInDb>("country_202302162148")
                select c;

            return listCountries.ToList();
        }

        //return an list of countries to replace as List<>
        public static List<CountriesToReplace> GetCountriesToReplace()
        {
            string actualPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var _excelFile = new ExcelQueryFactory(actualPath + "\\Excels\\new_countries.xlsx");

            //open excel file using linq to excel
            var file = _excelFile.GetWorksheetNames();

            //open spreadsheet
            var listCountries = from c in _excelFile.Worksheet<CountriesToReplace>("Feuil1")
                select c;

            return listCountries.ToList();
        }
    }
}
