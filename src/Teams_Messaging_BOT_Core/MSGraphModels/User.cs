using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.MSGraphModels
{
    public class UserResponseModel
    {
        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }
        public List<object> businessPhones { get; set; }
        public string displayName { get; set; }
        public string givenName { get; set; }
        public object jobTitle { get; set; }
        public string mail { get; set; }
        public object mobilePhone { get; set; }
        public object officeLocation { get; set; }
        public object preferredLanguage { get; set; }
        public string surname { get; set; }
        public string userPrincipalName { get; set; }
        public string id { get; set; }
    }

    public class TagsListResponseModel
    {
        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }

        [JsonProperty("@odata.count")]
        public int OdataCount { get; set; }
        public List<TagsResponseModel> value { get; set; }
    }
    public class TagsResponseModel
    {
        [JsonProperty("@odata.type")]
        public string OdataType { get; set; }
        public string id { get; set; }
        public string teamId { get; set; }
        public string displayName { get; set; }
        public string description { get; set; }
        public int memberCount { get; set; }
        public string tagType { get; set; }
    }
}
