using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace ToolBoxUtility
{
    // Class DistributionGroup
    [DataContract(Name = "distribution-group")]
    [XmlRoot("distribution-group")]
    public class DistributionGroup
    {
        public DistributionGroup()
        {
            error = "";
        }

        [XmlElement("id")]
        public string Alias { get; set; }

        [DataMember(Name = "name")]
        [XmlElement("name")]
        public string Name { get; set; }
        [DataMember(Name = "displayName")]
        [XmlElement("displayName")]
        public string displayName { get; set; }
        [DataMember(Name = "groupType")]
        [XmlElement("groupType")]
        public string groupType { get; set; }
        [DataMember(Name = "primarySmtpAddress")]
        [XmlElement("primarySmtpAddress")]
        public string primarySmtpAddress { get; set; }
        [DataMember(Name = "error")]
        [XmlElement("error")]
        public string error { get; set; }

        [DataMember(Name = "exchange-users")]
        [XmlElement("ExchangeUsers")]
        public ExchangeUserMembers users { get; set; }

        [XmlElement("ou")]
        public string OrganizationalUnit { get; set; }

        [XmlElement("has_children")]
        public bool HasChildren { get; set; }
    }

    [XmlRoot("distribution-groups")]
    public class DistributionGroupsShorter
    {
        [XmlAttribute("current_page")]
        public string CurrentPageDisplay
        {
            get { return CurrentPage.HasValue && CurrentPage.Value > 0 ? CurrentPage.Value.ToString() : null; }
            set
            {
                int val;
                if (int.TryParse(value, out val))
                {
                    CurrentPage = val;
                }
                else
                {
                    CurrentPage = null;
                }
            }
        }

        [XmlIgnore]
        public int? CurrentPage { get; set; }
        
        [XmlAttribute("per_page")]
        public string PerPageDisplay
        {
            get { return PerPage.HasValue && PerPage.Value > 0 ? PerPage.Value.ToString() : null; }
            set
            {
                int val;
                if (int.TryParse(value, out val))
                {
                    PerPage = val;
                }
                else
                {
                    PerPage = null;
                }
            }
        }

        [XmlIgnore]
        public int? PerPage { get; set; }

        [XmlAttribute("total_entries")]
        public int TotalEntries { get; set; }

        [XmlAttribute("type")]
        public string type = "array";

        [XmlElement("distribution-group")]
        public List<DistributionGroup> groups { get; set; }
    }

    [XmlRoot("DstrSvc")]
    public class DistributionGroupCreationParams
    {
        [XmlElement("group-name")]
        public string Name { get; set; }

        [XmlElement("ou")]
        public string OrganizationalUnit { get; set; }
    }
}