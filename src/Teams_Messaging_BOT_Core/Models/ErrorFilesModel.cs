using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.Models
{
    public class ErrorFilesModel
    {
        public DateTime RecordDate { get; set; }
        public string FileName { get; set; }
        public string ErroDetails { get; set; }
        public DateTime LastRetryDate { get; set; }
        public int RetryCount { get; set; }

    }
}
