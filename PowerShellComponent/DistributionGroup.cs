using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ToolBoxUtility;

namespace PowerShellComponent
{
    [Serializable()]
    public class DistributionGroupBak
    {
        public string name { get; set; }
        public string displayName { get; set; }
        public string groupType { get; set; }
        public string primarySmtpAddress { get; set; }
        public string error { get; set; }
        public List<ExchangeUser> users {get; set;}

        // ExchangeUser()
        // desc: Constructor
        public DistributionGroupBak()
        {
            name               = "";
            displayName        = "";
            groupType          = "";
            primarySmtpAddress = "";
            error              = "";
         //   users              = null;
        }
    }
}
