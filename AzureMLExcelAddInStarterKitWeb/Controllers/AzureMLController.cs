using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using System.Xml;
using System.Web.Services.Protocols;

using System.Net.Http.Formatting;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Newtonsoft.Json;


namespace AzureMLExcelAddInWeb.Controllers
{
    public class AzureMLController : ApiController
    {

        [HttpGet()]
        public string RunAzureML(string URLText, string APIKey, string MLColumns, string  MLRows)
        {
            try
            {

                InvokeRequestResponseService( URLText,  APIKey,  MLColumns,  MLRows).Wait();
                return this.MLResult;


            }
            catch (Exception e)
            {
                return "There was an exception: " + e.Message + "\n\n" + e.StackTrace;
            }
        }

        private string MLResult;

        public class Column
        {
            public string Name { get; set; }
        }



        public async Task InvokeRequestResponseService(string URLText, string APIKey, string MLColumns, string MLRows)
        {

            using (var client = new HttpClient())
            {
                JavaScriptSerializer js = new JavaScriptSerializer();
                string[] columns = js.Deserialize<string[]>(MLColumns);

                string[,] rows = JsonConvert.DeserializeObject<string[,]>(MLRows);

                var scoreRequest = new
                {

                    Inputs = new Dictionary<string, StringTable>() {
                        {
                            "input1",
                            new StringTable()
                            {
                                //ColumnNames = new string[] {"age", "workclass", "fnlwgt", "education", "education-num", "marital-status", "occupation", "relationship", "race", "sex", "capital-gain", "capital-loss", "hours-per-week", "native-country", "income"},
                                ColumnNames =columns,

                                //Values = new string[,] {  { "40", "Private", "155594", "Masters", "9", "Married-civ-spouse", "Sales", "Husband", "White", "Male", "0", "0", "40", "United-States", "<=50K" }, { "80", "Private", "331474", "Masters", "9", "Married-civ-spouse", "Adm-clerical", "Wife", "White", "Female", "0", "0", "40", "United-States", "<=50K" }, }
                                Values = rows
                            }
                        },
                    },
                    GlobalParameters = new Dictionary<string, string>()
                    {
                    }
                };
                //Testing
                //APIKey = "D1N/VNr+2BYidVeuH+2rSmFHCuid4QHw4MoLgrV7cxKAsG2spnOrjByRNGaE5oTIt4RohlSL5O/I26+UeKc1mw=="; // Replace this with the API key for the web service

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", APIKey);
                //Testing
                //client.BaseAddress = new Uri("https://ussouthcentral.services.azureml.net/workspaces/f8e2d18f739148248b1c73fbfae07fe2/services/28070a209e724b0c9a1b634c32d287f4/execute?api-version=2.0&details=true");
                client.BaseAddress = new Uri(URLText);

                // WARNING: The 'await' statement below can result in a deadlock if you are calling this code from the UI thread of an ASP.Net application.
                // One way to address this would be to call ConfigureAwait(false) so that the execution does not attempt to resume on the original context.
                // For instance, replace code such as:
                //      result = await DoSomeTask()
                // with the following:
                //      result = await DoSomeTask().ConfigureAwait(false)


                HttpResponseMessage response = await client.PostAsJsonAsync("", scoreRequest).ConfigureAwait(false);

                if (response.IsSuccessStatusCode)
                {
                    this.MLResult = await response.Content.ReadAsStringAsync();

                }
                else
                {
                    this.MLResult = await response.Content.ReadAsStringAsync();

                    // Get the headers - they include the requert ID and the timestamp, which are useful for debugging the failure
                    //return string.Format("The request failed with status code: {0}", response.StatusCode) + response.Headers.ToString() + responseContent;
                }
            }
        }
    }
     public class StringTable
    {
        public string[] ColumnNames { get; set; }
        public string[,] Values { get; set; }
    }

}