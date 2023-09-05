using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.MSGraphModels
{
    public class ChannelListResponseModel
    {
        public List<TeamValueResponse> value { get; set; }
    }
    public class ChannelValueResponse
    {
        public string id { get; set; }
        public DateTime createdDateTime { get; set; }
        public string displayName { get; set; }
        public string description { get; set; }
        public string membershipType { get; set; }
    }
}
