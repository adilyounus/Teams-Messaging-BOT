using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.Models
{
    public enum MessageToType
    {
        Chat = 1,
        Channel = 2
    }
    public class MessageInfoModel
    {
        public MessageToType messageToType { get; set; }
        public string To { get; set; }
        public string Title { get; set; }
        public string SubTitle { get; set; }
        public string Message { get; set; }
        public bool HasImage { get; set; }
        public string ImagePath { get; set; }
        public bool HasFile { get; set; }
        public string FilePath { get; set; }
    }
}
