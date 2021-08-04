using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public delegate void MessageEventHandler(object sender, MessageEventArgs e);
    public class MessageEventArgs : EventArgs
    {
        public string Id { get; set; }
        public string Message { get; set; }
    }
}