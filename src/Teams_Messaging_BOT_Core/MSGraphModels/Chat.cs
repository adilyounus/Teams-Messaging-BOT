using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.MSGraphModels
{
    public class Member
    {
        [JsonProperty("@odata.type")]
        public string OdataType { get; set; }
        public string id { get; set; }
        public List<object> roles { get; set; }
        public string displayName { get; set; }
        public string userId { get; set; }
        public string email { get; set; }
    }

    public class ChatListResponseModel
    {
        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }

        [JsonProperty("@odata.count")]
        public int OdataCount { get; set; }
        public List<ChatValueResponse> value { get; set; }
    }

    public class ChatValueResponse
    {
        public string id { get; set; }
        public string topic { get; set; }
        public DateTime createdDateTime { get; set; }
        public DateTime lastUpdatedDateTime { get; set; }
        public string chatType { get; set; }
        public List<Member> members { get; set; }
    }
}
