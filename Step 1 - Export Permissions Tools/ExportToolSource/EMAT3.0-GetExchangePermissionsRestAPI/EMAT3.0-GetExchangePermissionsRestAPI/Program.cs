using System;
using System.Net;

using System.Net.Http;
using System.Threading.Tasks;

namespace EMAT3._0_GetExchangePermissionsRestAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadContacts().Wait();
        }

        private static async Task ReadContacts()
        {
            var handler = new HttpClientHandler();
            handler.Credentials = new NetworkCredential()
            {
                UserName = ConfigurationManager.AppSettings["UserName"],
                Password = ConfigurationManager.AppSettings["Password"]
            };

            using (var client = new HttpClient(handler))
            {
                var url = "https://outlook.office365.com/api/v1.0/me/contacts";
                var result = await client.GetStringAsync(url);
                

                /*
                var data = JObject.Parse(result);

                foreach (var item in data["value"])
                {
                    Console.WriteLine(item["DisplayName"]);
                }*/
            }
        }


    }
}