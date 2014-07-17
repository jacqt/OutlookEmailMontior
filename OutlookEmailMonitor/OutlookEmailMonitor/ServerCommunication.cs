using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;

namespace OutlookEmailMonitor
{
    class ServerCommunication
    {
        String server_url;
        public ServerCommunication(String _server_url)
        {
            server_url = _server_url;
        }

        public async Task<String> performPostRequest(List<Parameter> parameters)
        {
            using (WebClient wb = new WebClient())
            {
                NameValueCollection data = new NameValueCollection();
                foreach (Parameter parameter in parameters)
                {
                    data[parameter.Item1] = parameter.Item2;
                }
                try
                {
                    byte[] response = wb.UploadValues(server_url, "POST", data);
                    String str_response = System.Text.Encoding.Default.GetString(response);
                    return str_response;
                }
                catch (Exception e)
                {
                    return "Failed...?";
                }
            }
        }

        //private async Task<String> post(String url)
        //{
        //    //System.Net.WebRequest;
        //    HttpWebRequest post_request = (HttpWebRequest)WebRequest.Create(url);
        //    post_request.Method = "POST";
        //    WebResponse response = post_request.GetResponse();
        //    Stream response_stream = response.GetResponseStream();
        //    StreamReader reader = new StreamReader(response_stream);
        //    String response_string = reader.ReadToEnd();
        //    return response_string;
        //}
    }
}