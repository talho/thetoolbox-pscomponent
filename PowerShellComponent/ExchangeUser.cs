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

// Scope TALHO
namespace ToolBoxUtility
{
    // Class ExchangeUser
    [DataContract(Name="exchange-user")]
    [XmlRoot("exchange-user")]
    public class ExchangeUser
    {
        public ExchangeUser() { error = ""; }

        // Declare public fields
        [XmlElement("identity")]
        public string identity { get; set; }
        [XmlElement("password")]
        public string password { get; set; }

        [DataMember(Name = "alias")]
        [XmlElement("alias")]
        public string alias { get; set; }
        [DataMember(Name = "dn")]
        [XmlElement("dn")]
        public string dn { get; set; }
        [DataMember(Name = "cn")]
        [XmlElement("cn")]
        public string cn { get; set; }
        [DataMember(Name = "upn")]
        [XmlElement("upn")]
        public string upn { get; set; }
        [DataMember(Name = "mailboxEnabled")]
        [XmlElement("mailboxEnabled")]
        public bool mailboxEnabled { get; set; }
        [DataMember(Name = "ou")]
        [XmlElement("ou")]
        public string ou { get; set; }
        [DataMember(Name = "login")]
        [XmlElement("login")]
        public string login { get; set; }
        [DataMember(Name = "email")]
        [XmlElement("email")]
        public string email { get; set; }
        [DataMember(Name = "has_vpn")]
        [XmlElement("has-vpn")]
        public bool has_vpn { get; set; }
        [DataMember(Name = "error")]
        [XmlElement("error")]
        public string error { get; set; }
        [XmlElement("type")]
        public string type { get; set; }
    }

    // Class ExchangeUserCollection, can take in a list of ExchangeUser, serialize them, then return an XML with a list of ExchangeUsers
    public class ExchangeUserCollection : List<ExchangeUser>
    {
        public int current_page { get; set; }
        public int per_page { get; set; }
        public int total_entries { get; set; }

        // toXML()
        // desc: Method serializes a list of ExchangeUsers, returns serialized XML
        // params: none
        // return: ExchangeUsers, IXmlSerialized
        public ExchangeUsers toXML()
        {
            DataContractSerializer serializer = new DataContractSerializer(typeof(ExchangeUser));
            using (MemoryStream memoryStream = new MemoryStream())
            {
                foreach (ExchangeUser user in this) serializer.WriteObject(memoryStream, user);
                memoryStream.Position = 0;

                String xml = null;
                using (StreamReader reader = new StreamReader(memoryStream))
                {
                    xml = reader.ReadToEnd();
                    reader.Close();
                }
                memoryStream.Close();
                return new ExchangeUsers(xml, total_entries, current_page, per_page);
            }
        }
    }

    // Class ExchangeUsers, takes in an XML object of ExchangeUser
    [XmlRoot(ElementName = "exchange-users")]
    public class ExchangeUsers : IXmlSerializable
    {
        public ExchangeUsers() { size = 1; current_page = 1; per_page = 1; }

        private int size;
        private int current_page;
        private int per_page;

        // ExchangeUsers()
        // desc: Constructor, set Content, size, current_page, and per_page
        // params: string content   - XML string 
        //         int size         - total number of entries
        //         int current_page - page to return
        //         int per_page     - number of entries to return per page
        // return: none
        public ExchangeUsers(string content, int size, int current_page, int per_page)
        {
            this.Content      = content;
            this.size         = size;
            this.current_page = current_page;
            this.per_page     = per_page;
        }

        public string Content { get; set; }

        //Overload IXmlSerializable
        #region IXmlSerializable Members

        public System.Xml.Schema.XmlSchema GetSchema()
        {
            return null;
        }

        public void ReadXml(XmlReader reader)
        {
            throw new NotImplementedException();
        }

        public void WriteXml(XmlWriter writer)
        {
            writer.WriteAttributeString("type", "array");
            writer.WriteAttributeString("current_page", current_page.ToString());
            writer.WriteAttributeString("per_page", per_page.ToString());
            writer.WriteAttributeString("total_entries", size.ToString());
            writer.WriteRaw(this.Content);
        }
        #endregion
    }

    [XmlRoot("exchange-users")]
    public class ExchangeUserShorter
    {
        [XmlAttribute("current_page")]
        public int CurrentPage { get; set; }
        [XmlAttribute("per_page")]
        public int PerPage { get; set; }
        [XmlAttribute("total_entries")]
        public int TotalEntries { get; set; }

        [XmlAttribute("type")]
        public string type = "array";

        [XmlElement("exchange-user")]
        public List<ExchangeUser> users { get; set; }
    }

    [XmlRoot("ExchangeUsers")]
    public class ExchangeUserMembers
    {
        [XmlAttribute("type")]
        public string type = "array";

        [XmlElement("ExchangeUser")]
        public List<ExchangeUser> users { get; set; }
    }
}
