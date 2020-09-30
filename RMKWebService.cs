using System.ServiceModel;

namespace RmkTwoCaPService
{

    // 
    // Этот исходный код был создан с помощью wsdl, версия=4.6.1055.0.
    // 


    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "4.6.1055.0")]
    [System.Web.Services.WebServiceBindingAttribute(Name = "BasicHttpBinding_CaPService", Namespace = "http://www.esterdev.com")]
    [ServiceContract(Name = "RmkTwoCaPService.IBasicHttpBinding_CaPService", Namespace = "http://www.esterdev.com")]
    public interface IBasicHttpBinding_CaPService
    {
        /// <remarks/>[System.Web.Services.WebMethodAttribute()]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.esterdev.com/CaPService/CreateCaPOperation", RequestNamespace = "http://www.esterdev.com", ResponseNamespace = "http://www.esterdev.com", Use = System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [System.ServiceModel.OperationContractAttribute(Action = "http://www.esterdev.com/CaPService/CreateCaPOperation", ReplyAction = "http://www.esterdev.com/CaPService/CreateCaPOperationResponse")]
        [return: System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
        string CreateCaPOperation([System.Xml.Serialization.XmlElementAttribute(IsNullable = true)] string data);
    }
}
