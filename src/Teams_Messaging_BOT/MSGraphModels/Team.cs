using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.MSGraphModels
{
    public class TeamListResponseModel
    {
        public List<TeamValueResponse> value { get; set; }
    }
    public class TeamValueResponse
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public string description { get; set; }
    }
}
