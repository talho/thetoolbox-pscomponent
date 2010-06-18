using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace PowerShellComponent
{
    [Serializable()]
    public class ExchangeUser
    {
        public string givenName { get; set; }
        public string sn { get; set; }
        public string dn { get; set; }
        public string cn { get; set; }
        public string mailbox { get; set; }
        public string alias { get; set; }
        public string upn { get; set; }
        public bool mailboxEnabled { get; set; }

        public ExchangeUser()
        {
            givenName = "";
            sn = "";
            dn = "";
            cn = "";
            mailbox = "";
            alias = "";
            upn = "";
            mailboxEnabled = false;
        }
    }
}
