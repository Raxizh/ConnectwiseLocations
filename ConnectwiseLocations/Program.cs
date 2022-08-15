using Newtonsoft.Json;
using QuickType;
using QuickType2;
using System.Text;
using System.Diagnostics;
using Aspose.Cells;
using System.Runtime.InteropServices;
using System.Reflection;

namespace ConnectwiseLocations
{
    public class Program
    {
        public static Stopwatch s = Stopwatch.StartNew();
        static void Main(string[] args)
        {
            s.Start(); //Timer
            var Companies = makeCompanySet();
            WriteCSV(Companies);
            
        }


        public static string WebRequest(string url)
        {
            using (var client = new HttpClient())
            {
                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url))
                {
                    try
                    {
                        client.DefaultRequestHeaders.Add("Authorization", "Basic Z2RzK1p1MmIxRHI4UHJKTk1pR0Q6b2lMMG40a3MzYkgySUJndg=="); //JacobB authorization added -- 7/12/22
                        client.DefaultRequestHeaders.Add("clientID", "be97d76e-6436-4cd7-ab32-9e8e86369453"); //JacobB clientID added -- 7/12/22
                        var response = client.SendAsync(request).Result;
                        var content = response.Content.ReadAsStringAsync().Result;

                        if (!response.IsSuccessStatusCode)
                        {
                            throw new Exception($"{content}: {response.StatusCode}");
                        }
                        return content;
                    }
                    catch
                    {
                        Console.WriteLine("Your Authorization or ClientID is invalid, please try again.");
                        Environment.Exit(0);
                    }
                }
                return "Invalid"; // Returns invalid if exception is caught
            }
        }


        // Web request for generating what the new site name should be
        public static string MiddlewareWebRequest(Site a)

        {
            string a1 = a.AddressLine1;
            string a2 = a.AddressLine2;
            string city = a.City;
            string state = a.Company.State;
            string zipCode = a.Zip;


            var url = "http://address.middleware1.getgds.com/Home/AuthenticateAddress?address=" + a1 + ", " + a2 + ", " + city + ", " + state + ", " + zipCode;

            using (var client = new HttpClient())
            {
                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url))
                {

                    var settings = new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    };

                    var response = client.SendAsync(request).Result;
                    var content = response.Content.ReadAsStringAsync().Result;
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception($"{content}: {response.StatusCode}");
                    }
                    try
                    {
                        var results = JsonConvert.DeserializeObject<AddressClass>(content, settings);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        if (results.ValidatedAddress != null && results.ValidatedAddress.SublocationId != null)
                        {
                            return results.ValidatedAddress.LocationId + "." + results.ValidatedAddress.SublocationId;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        }
                        else return null;
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
        }


        // Web request for getting a list of companies
        public static List<Company> getCompanies(int pageSize, int pageNum)
        {
            var settings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };

            var url = "https://connectwise.getgds.com/v4_6_release/apis/3.0/company/companies/?conditions=deletedflag=false&pagesize=" +  pageSize + "&page=" + pageNum; //deletedFlag = true to get a list of deleted sites(Do the same in the getCount function)
            var results = JsonConvert.DeserializeObject<List<Company>>(WebRequest(url), settings);
            return results;
        }


        // Gets a list of sites that a company has
        public static List<Site> getSites(long id, int pageSize, int pageNum)
        {
            var settings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };

            var url = "https://connectwise.getgds.com/v4_6_release/apis/3.0/company/companies/" + id + "/sites?pageSize=" + pageSize + "&page=" + pageNum;
            var results = JsonConvert.DeserializeObject<List<Site>>(WebRequest(url), settings);
            return results;
        }


        // Returns number of companies connectwise has
        public static long getCount()
        {
            var settings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };

            var url = "https://connectwise.getgds.com/v4_6_release/apis/3.0/company/companies/count/?conditions=deletedflag=false"; // deletedFlag = true to get a list of deleted sites (Do the same in the getCompanies function)
            var results = JsonConvert.DeserializeObject<Number>(WebRequest(url), settings);
            return results.Count;
        }


        // Checks if the current name for the site is valid
        public static bool isValidName(Site s)
        {
            string currName = s.Name;
            char[] delimiterChars = { ' ', '|', '.' };
            string[] words = currName.Split(delimiterChars);

            if (words.Length > 2)
            {
                string delimitedName = words[0] + "." + words[1] + "." + words[2];

                if (delimitedName == s.newSiteName)
                {
                    return true;
                }
                else return false;
            }
            return false;
        }


        // Checks if sites have the same new site name to flag them as duplicates, then picks the most recently updated as the site to keep 
        public static void dupSites(List<Site> l)
        {
            var dupList = l.GroupBy(s => s.newSiteName).Select(grp => grp.ToList());
            foreach (var i in dupList)
            {
                if (i.Count() > 1)
                {
                    var tempList = new List<Site>();
                    foreach (var j in i)
                    {
                        if (j.newSiteName != "")
                        {
                            j.isDuplicate = true;
                            tempList.Add(j);
                        }
                    }
                    if (tempList.Count > 1)
                    {
                        // Determines the main site based on if it is flagged as primary address or it is the last updated site in that order
                        var latestUpdate = tempList.Max(s => s.Info.LastUpdated);
                        var bestSite = tempList.Find(s => s.Info.LastUpdated == latestUpdate);
                        var priSite = tempList.Find(s => s.PrimaryAddressFlag == true);
                        foreach (var s in tempList)
                        {
                            if (s.PrimaryAddressFlag == false && priSite != null)
                            {
                                s.DuplicateOf = priSite.Id;
                            }

                            else if (s.PrimaryAddressFlag == true || bestSite == s)
                            {
                                s.isDuplicate = false;
                            }

                            else
                            {
                                s.DuplicateOf = bestSite.Id;
                            }

                            Console.WriteLine(s.Id + " " + s.newSiteName + " " + s.DuplicateOf + " " + s.PrimaryAddressFlag);
                        }
                    }
                }
            }
        }


        // Makes a set of all companies. Adds list of sites, adds new site name, and checks for duplicate sites for each company
        public static HashSet<Company> makeCompanySet()
            {
            double LocationCount = getCount();
            int currPage = 1;
            double maxPageSize = 100; // Original -> 500 
            double x = LocationCount / maxPageSize; 
            double numPages = Math.Ceiling(x); 
            var CompanySet = new HashSet<Company>();
            var csv = new StringBuilder();

            while (currPage <= numPages) // Original -> (currPage <= numPages)
            {
                var temp = getCompanies((int)maxPageSize, currPage);
                foreach (var i in temp)

                {   
                    i.sitelist = getSites(i.Id, 1000, 1); // Original -> (i.Id, 1000, 1)
                    

                    foreach (var j in i.sitelist)
                    {
                        var charArray = j.Company.Identifier.ToCharArray();
                        if (j.AddressLine1 == null && j.AddressLine2 == null || j.City == null || charArray.Length > 6)
                        {
                            j.newSiteName = "";
                        }
                        else
                        {
                            string locSubloc = MiddlewareWebRequest(j);
                            if (locSubloc != null)
                            {
                                j.newSiteName = j.Company.Identifier + "." + locSubloc;
                            }
                            else
                            {
                                j.newSiteName = "";
                            }
                        }
                        j.isValid = isValidName(j);
                    }
                    
                    dupSites(i.sitelist);
                    CompanySet.Add(i);
                    
                }
                currPage++;
            }
            return CompanySet;
        }


        //Uses set of all companies to write the csv
        public static void WriteCSV(HashSet<Company> companySet)
        {            
            var csv = new StringBuilder();

            csv.AppendLine("SiteIDNumber," + "ID," + "Company," + "SiteName/duplicateID," + "address1," + "address2," + "City,"
                + "State,"+ "Zip," + "NewSiteName," + "isDuplicate," + "isValidName," + "DuplicateOfSite#," + "isDefaultBilling," + "isPrimaryAddress" + "companyId");
            foreach (var c in companySet)
            {
                foreach (var s in c.sitelist)
                {
                    string siteid = "\"" + s.Id + "\"";
                    string id = "\"" + s.Company.Identifier + "\"";
                    string Company = "\"" + s.Company.Name + "\"";
                    string SiteName = "\"" + s.Name + "\"";
                    string a1 = "\"" + s.AddressLine1 + "\"";
                    string a2 = "\"" + s.AddressLine2 + "\"";
                    string City = "\"" + s.City + "\"";
                    string State = "\"" + c.State + "\"";
                    string Zip = "\"" + s.Zip + "\"";
                    string newSiteName = "\"" + s.newSiteName + "\"";
                    string dup = "\"" + s.isDuplicate + "\"";
                    string valid = "\"" + s.isValid + "\"";
                    string dupOf = "\"" + s.DuplicateOf + "\"";
                    string billingFlag = "\"" + s.DefaultBillingFlag + "\"";
                    string primaryAddressBool = "\"" + s.PrimaryAddressFlag + "\"";
                    string compId = "\"" + s.Company.Id + "\"";

                    csv.AppendLine(siteid + "," + id + "," + Company + "," + SiteName + "," + a1 + "," + a2 + "," + City + ","
                        + State + "," + Zip + "," + newSiteName + "," + dup + "," + valid + "," + dupOf + "," + billingFlag + "," 
                        + primaryAddressBool + "," + compId);
                }
            }
            try
            {
                Console.WriteLine("Time taken: " + (s.ElapsedMilliseconds/1000) + " seconds"); // Timer
                s.Stop();

                //Checks to see if file already exists
                if (File.Exists("ConnectwiseSites.csv"))
                {
                    File.Delete("ConnectwiseSites.csv");
                }

                File.WriteAllText("ConnectwiseSites.csv", csv.ToString());

                var workbook = new Workbook("ConnectwiseSites.csv");

                //Checks to see if file already exists
                if (File.Exists("ConnectwiseSitesUpd.xlsm"))
                {
                    File.Delete("ConnectwiseSitesUpd.xlsm");
                }

                workbook.Save("ConnectwiseSitesUpd.xlsm");

                try
                {
                    RunMacro();
                }
                catch (Exception e) { Console.WriteLine(e.ToString()); }

            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }


        // Initiates an ExcelApplication variable and runs the Highlight_All_Dups1 macro created in the CreateWorkbook method
        public static void RunMacro()
        {
            string workingDirectory = Environment.CurrentDirectory;
            string filePath = workingDirectory + "\\ConnectwiseSitesUpd.xlsm";
            CreateWorkbook(filePath);
            Console.WriteLine("File saved to " + filePath);
            Console.WriteLine("Finished executing...");
            File.Delete("ConnectwiseSites.csv");

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Sheets worksheets = excelWorkbook.Worksheets;
            excelWorkbook.Application.Run("ConnectwiseSites_Macros.Highlight_All_Dups1", Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Missing.Value, Missing.Value);
            excelWorkbook.Save();
            excelWorkbook.Close();
            Marshal.ReleaseComObject(worksheets);
            excelApp.Quit();
        }


        // Creates the Macro-Enabled Excel Workbook as well as the Highlight_All_Dups Macro
        public static void CreateWorkbook(string fileName)
        {
            try
            {

                Workbook workbook = new Workbook(fileName);
                Worksheet worksheet = workbook.Worksheets[0];
                Aspose.Cells.Vba.VbaModule moduleTest = workbook.VbaProject.Modules[workbook.VbaProject.Modules.Add(worksheet)];
                moduleTest.Name = "ConnectwiseSites_Macros";
                moduleTest.Codes = "Sub Highlight_All_Dups1()" + "\r\n" +
                    "    Worksheets(\"ConnectwiseSites\").Activate" + "\r\n" +
                    "    Columns(\"K:K\").Select" + "\r\n" +
                    "    With Application.ReplaceFormat.Interior" + "\r\n" +
                    "        .PatternColorIndex = xlAutomatic" + "\r\n" +
                    "        .Color = 255" + "\r\n" +
                    "        .TintAndShade = 0" + "\r\n" +
                    "        .PatternTintAndShade = 0" + "\r\n" +
                    "    End With" + "\r\n" +
                    "    Selection.Replace What:=\"true\", Replacement:=\"TRUE\", LookAt:=xlPart, _" + "\r\n" +
                    "        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _" + "\r\n" +
                    "        ReplaceFormat:=True, FormulaVersion:=xlReplaceFormula2" + "\r\n" + "\r\n" +
                    "    Columns(\"C:C\").EntireColumn.AutoFit" + "\r\n" +
                    "    Columns(\"D:D\").EntireColumn.AutoFit" + "\r\n" +
                    "    ActiveWindow.ScrollColumn = 2" + "\r\n" +
                    "    ActiveWindow.ScrollColumn = 3" + "\r\n" +
                    "    ActiveWindow.ScrollColumn = 4" + "\r\n" +
                    "    ActiveWindow.ScrollColumn = 5" + "\r\n" +
                    "    Columns(\"J:J\").EntireColumn.AutoFit" + "\r\n" +
                    "    Columns(\"G:G\").EntireColumn.AutoFit" + "\r\n" +
                    "    Range(\"A1\").Select" + "\r\n" +
                    "    Sheets(\"Evaluation Warning\").Delete" + "\r\n" +
                    "    Sheets(\"Evaluation Warning (1)\").Delete" + "\r\n" +
                    "End Sub";
                workbook.Save(fileName, SaveFormat.Xlsm);
            }
            catch { }
            GC.Collect();
        }
    }
}