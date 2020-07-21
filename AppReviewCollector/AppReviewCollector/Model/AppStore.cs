using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace AppReviewCollector.Model
{
    [XmlRoot("feed", Namespace = "http://www.w3.org/2005/Atom")]
    public class AppStore
    {
        [XmlElement("id")]
        public string id { get; set; }

        [XmlElement("title")]
        public string title { get; set; }

        [XmlElement("updated")]
        public string updated { get; set; }


        [XmlElement("link")]
        public List<string> link { get; set; } = new List<string>();

        [XmlElement("icon")]
        public string icon { get; set; }

        [XmlElement("author")]
        public Author author { get; set; }


        [XmlElement("rights")]
        public string rights { get; set; }

        [XmlElement("entry")]
        public List<Entry> entry { get; set; }

    }

    public class Author
    {
        [XmlElement("name")]
        public string name;

        [XmlElement("uri")]
        public string uri;
    }

    public class Entry
    {
        [XmlElement("updated")]
        public string updated;

        [XmlElement("id")]
        public string id;

        [XmlElement("title")]
        public string title;

        [XmlElement("content")]
        public List<string> content;

        [XmlElement(Namespace = "http://itunes.apple.com/rss", ElementName = "contentType")]
        public string contentType;

        [XmlElement(Namespace = "http://itunes.apple.com/rss", ElementName = "voteSum")]
        public string voteSum;

        [XmlElement(Namespace = "http://itunes.apple.com/rss", ElementName = "voteCount")]
        public string voteCount;

        [XmlElement(Namespace = "http://itunes.apple.com/rss", ElementName = "rating")]
        public string rating;

        [XmlElement(Namespace = "http://itunes.apple.com/rss", ElementName = "version")]
        public string version;


        [XmlElement("author")]
        public Author author { get; set; }

        [XmlElement("link")]
        public string link;

//        [XmlElement("content")]
//        public string contentHtml;

    }
}
