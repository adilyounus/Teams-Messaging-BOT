using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.MSGraphModels
{
    public class APIErrorResponse
    {
        public APIErrorResponseDetail error { get; set; }
    }
    public class APIErrorResponseDetail
    {
        public string code { get; set; }
        public string message { get; set; }
        public APIErrorResponseInnerError innerError { get; set; }
    }

    public class APIErrorResponseInnerError
    {
        public DateTime date { get; set; }

        [JsonProperty("request-id")]
        public string RequestId { get; set; }

        [JsonProperty("client-request-id")]
        public string ClientRequestId { get; set; }
    }

    
}
