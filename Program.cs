using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Globalization;
using QuickType;
using QuickType2;
using System.Text;

namespace ConnectwiseLocations
{
    public class Program
    {
        static void Main(string[] args)
        {
            var Companies = makeCompanySet();
            WriteCSV(Companies);
        }
        public static string WebRequest(string url)
        { 
            using (var client = new HttpClient())
            {
                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url))
                {

                    client.DefaultRequestHeaders.Add("Authorization", "Basic Z2RzK0wwZkpuV2NZc3N5MVNuWE46aXNFY21kUVJaSGpXQTRRNQ==");
                    client.DefaultRequestHeaders.Add("clientID", "be97d76e-6436-4cd7-ab32-9e8e86369453");
                    var response = client.SendAsync(request).Result;
                    var content = response.Content.ReadAsStringAsync().Result;
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception($"{content}: {response.StatusCode}");
                    }
                    return content;
                }
            }
        }
        //Web request for generating what the new site name should be
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
        //Web request for getting a list of companies
        public static List<Company> getCompanies(int pageSize, int pageNum)
        {
            var settings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            var url = "https://connectwise.getgds.com/v4_6_release/apis/3.0/company/companies?conditions=deletedFlag=false&pageSize=" + pageSize + "&page=" + pageNum;
            var results = JsonConvert.DeserializeObject<List<Company>>(WebRequest(url), settings);
            return results;
        }
        //Gets a list of sites that a company has
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
        //Returns number of companies connectwise has
        public static long getCount()
        {
            var settings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };

            var url = "https://connectwise.getgds.com/v4_6_release/apis/3.0/company/companies/count?conditions=deletedFlag=false";
            var results = JsonConvert.DeserializeObject<Number>(WebRequest(url), settings);
            return results.Count;
        }
        /*
        public static string getTaxCode(List<Site> s, long id, int pageSize, int pageNum)
        {

            var settings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };

            var url = "https://connectwise.getgds.com/v4_6_release/apis/3.0/company/companies/" + id + "/sites?pageSize=" + pageSize + "&page=" + pageNum;
            var results = JsonConvert.DeserializeObject<string>(WebRequest(url), settings);
            return results;
        }
        */
        //Checks if the current name for the site is valid
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
        //Checks if sites have the same new site name to flag them as duplicates, then picks the most recently updated as the site to keep
        public static void dupSites(List<Site> l)
        {
            var dupList = l.GroupBy(s => s.newSiteName).Select(grp => grp.ToList());
            foreach (var i in dupList)
            {
                if(i.Count() > 1)
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

                        var latestUpdate = tempList.Max(s => s.Info.LastUpdated);
                        var bestSite = tempList.Find(s => s.Info.LastUpdated == latestUpdate);
                        var priSite = tempList.Find(s => s.PrimaryAddressFlag == true);
                        checkFlag(tempList);

                        foreach (var s in tempList)
                        {
                            if ( s.PrimaryAddressFlag == false && priSite != null)
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

                            Console.WriteLine(s.Id + " " + s.newSiteName + " " + s.DuplicateOf + " " 
                                            + s.PrimaryAddressFlag + " " + s.DefaultBillingFlag + " " 
                                            + s.DefaultMailingFlag + " " + s.DefaultShippingFlag); 
                        }
                    }
                }
            }
        }
        //Makes a set of all companies. Adds list of sites, adds new site name, and checks for duplicate sites for each company
        public static HashSet<Company> makeCompanySet()
        {
            double LocationCount = getCount();
            int currPage = 1;
            double maxPageSize = 500; //double maxPageSize = 500;
            double x = LocationCount / maxPageSize;
            double numPages = Math.Ceiling(x);
            var CompanySet = new HashSet<Company>();
            var csv = new StringBuilder();
    
            while (currPage < numPages) // while (currPage < numPages)
            {

                var temp = getCompanies((int)maxPageSize, currPage);

                foreach (var i in temp)
                {
                    i.sitelist = getSites(i.Id, 1000, 1); //getSites(i.Id, 1000, 1);
                    checkFlag(i.sitelist);
                    //var taxCode = getTaxCode(i.sitelist, i.Id, 10, 1);

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

        public static void checkFlag(List<Site> S)
        {

            foreach(Site s in S)
            {

                if (s.PrimaryAddressFlag)
                {
                    s.hasPrimFlag = true;
                }
                if (s.DefaultBillingFlag)
                {
                   s.hasBillFlag = true;
                }
                if (s.DefaultMailingFlag)
                {
                    s.hasMailFlag = true;
                }
                if (s.DefaultShippingFlag)
                {
                    s.hasShipFlag = true;
                }

            }
        }

        public static string ReplaceSiteName(Site s)
        {

            string siteName;
            string Replace;
            string currname = s.Name;
            string[] words = currname.Split('|');

            if (words.Length == 2 && s.newSiteName != "")
            {
                siteName = s.newSiteName + '|' + words[1];
                char[] charArr = siteName.ToCharArray();

                if (charArr.Length <= 50)
                {
                    return siteName;
                } 
                else 
                {
                    return siteName.Substring(0,50);
                }

            } 
            else if (words.Length == 2 && s.newSiteName == "")
            {

                siteName = words[0] + '|' + words[1];
                char[] charArr = siteName.ToCharArray();

                if (charArr.Length <= 50)
                {
                    return siteName;
                } 
                else
                {
                    return siteName.Substring(0, 50);
                }

            }
            else if (s.newSiteName == "")
            {
                
                char[] charArr = currname.ToCharArray();

                if (currname.Length <= 50)
                {
                    return currname;
                }
                else
                {
                    return currname.Substring(0, 50);
                }

            }
            else
            {

                Replace = s.newSiteName;
                char[] replaceArr = Replace.ToCharArray();

                if (replaceArr.Length <= 50)
                {
                    return Replace; 
                } 
                else
                {
                    return Replace.Substring(0, 50);
                }

            }
        }

        //Uses set of all companies to write the csv
        public static void WriteCSV(HashSet<Company> companySet)
        {
            var csv = new StringBuilder();
            csv.AppendLine("SiteIDNumber," + "ID," + "CompanyId," + "Company," + "SiteName," 
                            + "address1," + "address2," + "City,"
                            + "State,"+ "Zip," + "NewSiteName," + "isDuplicate," 
                            + "isValidName," + "DuplicateOfSite#," + "HasPrimFlag," 
                            + "HasBillFlag," + "HasShipFlag," + "hasMailFlag," + "ReplaceSiteName");

            foreach (var c in companySet)
            {
                foreach (var s in c.sitelist)
                {

                    string siteid = "\"" + s.Id + "\"";
                    string id = "\"" + s.Company.Identifier + "\"";
                    string companyId = "\"" + c.Id + "\"";
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
                    string hasPrimFlag = "\"" + s.hasPrimFlag + "\"";
                    string hasBillFlag = "\"" + s.hasBillFlag + "\"";
                    string hasShipFlag = "\"" + s.hasShipFlag + "\"";
                    string hasMailFlag = "\"" + s.hasMailFlag + "\"";
                    string Name = "\"" + ReplaceSiteName(s) + "\"";

                    csv.AppendLine(siteid + "," + id + "," + companyId + "," + Company + "," + SiteName + "," + a1 + "," 
                                    + a2 + "," + City + "," + State + "," + Zip + "," 
                                    + newSiteName + "," + dup + "," + valid + "," + dupOf + "," 
                                    + hasPrimFlag + "," + hasBillFlag 
                                    + "," + hasShipFlag + "," + hasMailFlag + "," + Name);
                }
            }
            File.WriteAllText("ConnectwiseSites.csv", csv.ToString());
        }
    }
}