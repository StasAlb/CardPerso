using System.Configuration;

namespace CardPerso
{
    public class MovingStockSection : ConfigurationSection
    {
        [ConfigurationProperty("fields")]
        public FieldsCollection Fields
        {
            get { return (FieldsCollection)this["fields"]; }
        }
    }

    public class FieldsCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new FieldElement();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((FieldElement)element).Product;
        }
    }
    public class FieldElement : ConfigurationElement
    {
        [ConfigurationProperty("product", IsRequired = true)]
        public string Product
        {
            get { return (string)this["product"]; }
            set { this["product"] = value; }
        }
        [ConfigurationProperty("columns")]
        public ColumnsCollection Columns
        {
            get { return (ColumnsCollection)this["columns"]; }
        }
    }
    public class ColumnsCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ColumnElement();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ColumnElement)element).Nums;
        }
    }
    public class ColumnElement : ConfigurationElement
    {
        [ConfigurationProperty("nums", IsRequired = true)]
        public string Nums
        {
            get { return (string)this["nums"]; }
            set { this["nums"] = value; }
        }
        [ConfigurationProperty("debet", IsRequired = true)]
        public string Debet
        {
            get { return (string)this["debet"]; }
            set { this["debet"] = value; }
        }
        [ConfigurationProperty("credit", IsRequired = true)]
        public string Credit
        {
            get { return (string)this["credit"]; }
            set { this["credit"] = value; }
        }
        [ConfigurationProperty("ground", IsRequired = true, DefaultValue = "")]
        public string Ground
        {
            get { return (string)this["ground"]; }
            set { this["ground"] = value; }
        }

    }

}
