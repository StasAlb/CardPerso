using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Web;

namespace CardPerso
{
    public class RabbitMQSection : ConfigurationSection
    {
        [ConfigurationProperty("enabled", IsRequired = false, DefaultValue = "true")]
        public string IsEnabled
        {
            get { return (string) this["enabled"]; }
            set { this["enabled"] = value; }
        }
        
        [ConfigurationProperty("endpoint", IsRequired = true)]
        public RabbitEndpoint rEndpoint
        {
            get { return (RabbitEndpoint)this["endpoint"]; }
            set { this["endpoint"] = value; }
        }
        [ConfigurationProperty("credential", IsRequired = true)]
        public RabbitCredential rCredential
        {
            get { return (RabbitCredential)this["credential"]; }
            set { this["credential"] = value; }
        }
        [ConfigurationProperty("channel", IsRequired = true)]
        public RabbitChannel rChannel
        {
            get { return (RabbitChannel)this["channel"]; }
            set { this["channel"] = value; }
        }

    }

    public class RabbitEndpoint : ConfigurationElement
    {
        [ConfigurationProperty("address", IsRequired = true)]
        public string RabbitAddress
        {
            get { return (string) this["address"]; }
            set { this["address"] = value; }
        }
        [ConfigurationProperty("host", IsRequired = true)]
        public string RabbitHost
        {
            get { return (string)this["host"]; }
            set { this["host"] = value; }
        }
    }
    public class RabbitCredential : ConfigurationElement
    {
        [ConfigurationProperty("login", IsRequired = true)]
        public string RabbitLogin
        {
            get { return (string)this["login"]; }
            set { this["login"] = value; }
        }
        [ConfigurationProperty("password", IsRequired = true)]
        public string RabbitPassword
        {
            get { return (string)this["password"]; }
            set { this["password"] = value; }
        }
    }
    public class RabbitChannel : ConfigurationElement
    {
        [ConfigurationProperty("exchange", IsRequired = true)]
        public string RabbitExchange
        {
            get { return (string)this["exchange"]; }
            set { this["exchange"] = value; }
        }
        [ConfigurationProperty("queue", IsRequired = true)]
        public string RabbitQueue
        {
            get { return (string)this["queue"]; }
            set { this["queue"] = value; }
        }
        [ConfigurationProperty("routingkey", IsRequired = true)]
        public string RabbitRoutingKey
        {
            get { return (string)this["routingkey"]; }
            set { this["routingkey"] = value; }
        }
    }
}