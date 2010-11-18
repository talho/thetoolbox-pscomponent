using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;

namespace ToolBoxUtility
{
    public static class XmlSerializationHelper
    {
        public static string Serialize<T>(T obj)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));

            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            serializer.Serialize(sw, obj);

            sw.Close();
            
            return sb.ToString();
        }

        public static T Deserialize<T>(string objString)
        {
            T obj = default(T);

            XmlSerializer serializer = new XmlSerializer(typeof(T));

            StringReader sr = new StringReader(objString);
            obj = (T)serializer.Deserialize(sr);

            sr.Close();

            return obj;
        }
    }
}
