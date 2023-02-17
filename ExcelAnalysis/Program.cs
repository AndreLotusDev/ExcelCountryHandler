using ExcelAnalysis.Helper;
using ExcelAnalysis.Models;

//ExtractDuplicates.Run();

var listCountriesInDb = ExtractDuplicates.GetCountriesInDb();

var listCountriesToReplace = ExtractDuplicates.GetCountriesToReplace();


var missingCountry =
    listCountriesInDb.Where(a => a.code != null && a.code != "" && !listCountriesToReplace.Any(replace => a.code?.Trim().ToLower() == replace.country_code?.Trim().ToLower()));

var havingCountry =
    listCountriesInDb.Where(a => a.code != null && a.code != "" && listCountriesToReplace.Any(replace => a.code?.Trim().ToLower() == replace.country_code?.Trim().ToLower()));

List<(CountriesInDb, CountriesToReplace)> countryTuple = new List<(CountriesInDb, CountriesToReplace)>();
//Insert in country tuple the same country in different list
foreach (var country in havingCountry.OrderBy(o => o.country_name))
{
    var countryToReplaceFound = listCountriesToReplace.FirstOrDefault(f => f.country_code?.ToLower().Trim() == country.code?.ToLower().Trim());
    countryTuple.Add((country, countryToReplaceFound));
}

List<string> sqlCommand = new();
List<string> sqlCommandRenameEntityAddress = new();

foreach (var tuple in countryTuple)
{
    Console.WriteLine(tuple.Item1.country_name + "<=>" + tuple.Item2.name);
    var tempCommand = $"UPDATE country SET country_name = '{tuple.Item2.name.Replace("'", "''")}' WHERE code = '{tuple.Item2.country_code.Replace("'", "''")}';";
    sqlCommand.Add(tempCommand);
}

foreach (var tuple in countryTuple)
{
    var tempCommand = $"UPDATE entity_address SET country_name = '{tuple.Item2.name.Replace("'", "''")}' WHERE country_name = '{tuple.Item1.country_name.Replace("'", "''")}';";
    sqlCommandRenameEntityAddress.Add(tempCommand);
}

//insert sql commands in different files in desktop with different name files
System.IO.File.WriteAllLines(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +"\\sqlCommand.txt", sqlCommand);
System.IO.File.WriteAllLines(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\sqlCommandRenameEntityAddress.txt", sqlCommandRenameEntityAddress);

//print all missing country
foreach (var country in missingCountry)
{
    var countryToReplaceFound = listCountriesToReplace.FirstOrDefault(f => f.country_code?.ToLower().Trim() == country.code?.ToLower().Trim());
    Console.WriteLine(countryToReplaceFound?.name + " " + countryToReplaceFound?.country_code);
    Console.WriteLine(country.country_name + " <=> " + country.code);
    //Write a wide liner
    Console.WriteLine("--------------------------------------------------------------------------------------------------------------------");
}
