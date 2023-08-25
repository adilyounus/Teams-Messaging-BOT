using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT
{
    public class RetryMessageException : Exception
    {
        public RetryMessageException(string Message) : base(Message)
        {

        }
    }

    public class WithoutNotifyException : Exception
    {
        public WithoutNotifyException(string Message) : base(Message)
        {

        }
    }

    public class WithoutRetryException : Exception
    {
        public WithoutRetryException(string Message) : base(Message)
        {

        }
    }
}
