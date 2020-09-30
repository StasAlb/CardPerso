using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace CardPerso
{
    public class Products
    {
        public static SmsInfoSection _section = ConfigurationManager.GetSection("sms_info") as SmsInfoSection;

        public static SmsInfoElementCollection GetProducts()
        {
            return _section.SmsInfo;
        }
    }

    //Extend the ConfigurationSection class.  Your class name should match your section name and be postfixed with "Section".
    public class SmsInfoSection : ConfigurationSection
    {
        //Decorate the property with the tag for your collection.
        [ConfigurationProperty("products")]
        public SmsInfoElementCollection SmsInfo
        {
            get { return (SmsInfoElementCollection) this["products"]; }
        }
    }

    //Extend the ConfigurationElementCollection class.
    //Decorate the class with the class that represents a single element in the collection.
    [ConfigurationCollection(typeof(SmsInfoElement))]
    public class SmsInfoElementCollection : ConfigurationElementCollection
    {
        [ConfigurationProperty("CompanyField", IsRequired = true)]
        public string CompanyField
        {
            get { return (string)this["CompanyField"]; }
            set { this["CompanyField"] = value; }
        }

        public SmsInfoElement this[int index]
        {
            get { return (SmsInfoElement) BaseGet(index); }
            set
            {
                if (BaseGet(index) != null)
                    BaseRemoveAt(index);
                BaseAdd(index, value);
            }
        }

        protected override ConfigurationElement CreateNewElement()
        {
            return new SmsInfoElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((SmsInfoElement) element).ShortName;
        }
    }

    //Extend the ConfigurationElement class.  This class represents a single element in the collection.
    //Create a property for each xml attribute in your element.
    //Decorate each property with the ConfigurationProperty decorator.  See MSDN for all available options.
    public class SmsInfoElement : ConfigurationElement
    {
        [ConfigurationProperty("shortname", IsRequired = true)]
        public string ShortName
        {
            get { return (string) this["shortname"]; }
            set { this["shortname"] = value; }
        }

        [ConfigurationProperty("bin", IsRequired = true)]
        public string Bin
        {
            get { return (string) this["bin"]; }
            set { this["bin"] = value; }
        }

        [ConfigurationProperty("code", IsRequired = true)]
        public string Code
        {
            get { return (string) this["code"]; }
            set { this["code"] = value; }
        }

        [ConfigurationProperty("allcards", IsRequired = false, DefaultValue = "false")]
        public string AllCards
        {
            get { return (string) this["allcards"]; }
            set { this["allcards"] = value; }
        }
    }
}