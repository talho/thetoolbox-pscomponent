using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Channels;
using System.Xml;
using System.ServiceModel;

namespace ToolBoxUtility
{
    public static class MessageBuilder
    {
        private class BodyBuilder : BodyWriter
        {
            XmlElement xmlElement;

            public BodyBuilder(XmlElement xmlElement)
                : base(true)
            {
                this.xmlElement = xmlElement;
            }

            protected override void OnWriteBodyContents(XmlDictionaryWriter writer)
            {
                xmlElement.RemoveAttribute("xmlns:xsi");
                xmlElement.RemoveAttribute("xmlns:xsd");
                xmlElement.WriteTo(writer);
            }
        }

        public static Message CreateResponseMessage(XmlDocument inputDocument)
        {
            BodyBuilder writer = new BodyBuilder(inputDocument.DocumentElement);

            Message msg = Message.CreateMessage(MessageVersion.None,
                OperationContext.Current.OutgoingMessageHeaders.Action, writer);

            return msg;
        }


        public static Message CreateResponseMessage<T>(T obj)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(XmlSerializationHelper.Serialize(obj));
            BodyBuilder writer = new BodyBuilder(doc.DocumentElement);
            return Message.CreateMessage(MessageVersion.None, OperationContext.Current.OutgoingMessageHeaders.Action, writer);
        }
    }
}
