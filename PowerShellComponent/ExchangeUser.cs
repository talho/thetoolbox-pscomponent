using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

// Scope PowerShellComponent
namespace PowerShellComponent
{
    // Class ExchangeUser
    [Serializable()]
    public class ExchangeUser
    {
        public string alias { get; set; }
        public string dn { get; set; }
        public string cn { get; set; }
        public string upn { get; set; }
        public bool mailboxEnabled { get; set; }
        public string ou { get; set; }
        public string login { get; set; }
        public string email { get; set; }
        public bool has_vpn { get; set; }

        // ExchangeUser()
        // desc: Constructor
        public ExchangeUser()
        {
            dn             = "";
            cn             = "";
            upn            = "";
            ou             = "";
            login          = "";
            email          = "";
            has_vpn        = false;
            mailboxEnabled = false;
        }
    }
}
