<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Administration/Administration.Master" CodeBehind="EMail.aspx.cs" Inherits="CardPerso.Administration.EMail" %>
<asp:Content ID="ManageUserContent" runat="server" ContentPlaceHolderID="AdministrationContent">
    Кому: <asp:TextBox ID="tbTo" runat="server"></asp:TextBox>
<asp:Button ID="bSend" runat="server" onclick="bSend_Click" Text="Послать" />
</asp:Content>